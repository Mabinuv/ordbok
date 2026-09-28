/**
 * Ordbok – Swedish vocabulary backend (Google Sheets + LLM auto-fill)
 *
 * SETUP
 * 1. Create a Google Sheet. Extensions > Apps Script. Paste this whole file, save.
 * 2. Project Settings (gear) > Script properties > add the key for your provider:
 *      OPENAI_API_KEY = sk-...            (when PROVIDER = 'openai')
 *      GEMINI_API_KEY = ...               (when PROVIDER = 'gemini')
 * 3. OPTIONAL — shared PIN, worth setting once this is on a public URL: add
 *      APP_PIN = 1234                     (any digits/letters you like)
 *    Every save, edit and auto-fill request must then include this PIN; search
 *    stays open to anyone with the link. Leave APP_PIN unset to require no PIN
 *    at all (fine for a private/personal deployment).
 * 4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *    (Later code changes: Deploy > Manage deployments > pencil > Version: New version.)
 * 5. Paste the /exec URL into the web app's "Connect Google Sheet" panel.
 *
 * The "Words" tab is created on first use. Any column below that's missing from
 * an older sheet (Learned, Grammar Forms, ...) is added automatically at the end
 * — columns are matched by header NAME, not position, so this is safe to re-run.
 */

const SHEET_NAME = 'Words';
// Column order for a brand-new sheet. On an existing sheet, columns are looked
// up by name (see headerMap_), so this list only decides where NEW columns land.
const HEADERS = [
  'Swedish Word',
  'Pronunciation Guide',
  'Part of Speech',
  'English Meaning',
  'Example Sentence (Swedish)',
  'Example Sentence (English)',
  'Learned',
  'Grammar Forms',      // JSON: [{label, sv, example_sv, example_en}, ...]
  'Seen As',            // JSON: ["gick", ...] — inflected forms the user actually typed/looked up
  'Review State',       // JSON: {ef, interval, reps, due, last, lapses} — SM-2-lite spaced-repetition state
  'Tags',               // comma-separated, e.g. "food, home" — meant to be hand-edited in the Sheet too
  'Source',             // free text: where/how the word was encountered (also grounds the AI example sentence)
];

// ---- Auto-fill settings: check these two match what you use ----
const PROVIDER = 'openai';                       // 'openai' or 'gemini'
const OPENAI_MODEL = 'gpt-5.4-mini';
const GEMINI_MODEL = 'gemini-3.5-flash-lite';
const POS_OPTIONS = ['noun (en)', 'noun (ett)', 'verb', 'adjective', 'adverb', 'pronoun',
  'preposition', 'conjunction', 'interjection', 'numeral', 'phrase'];

// Hard cap on generated tokens. Keeps every auto-fill call short (fields are
// meant to be quick reference notes, not essays) and cuts response time a lot,
// since latency scales mostly with how many tokens the model has to generate.
const AUTOFILL_MAX_TOKENS = 450;

// Same rationale as AUTOFILL_MAX_TOKENS, sized for a 100-150 word Swedish
// passage plus the wordsUsed array and JSON overhead.
const READING_MAX_TOKENS = 700;
const READING_WORD_COUNT = { min: 100, max: 150 };

// How long a search result set is cached (seconds). Search re-reads the whole
// sheet otherwise, and a Sheets API read is the slowest part of a search by
// far, so this is the main lever for making search feel instant. Any
// save/edit/learn immediately clears the cache so you never see stale data.
const SEARCH_CACHE_SECONDS = 50;

function autofill_(word, sourceSentence) {
  const propName = PROVIDER === 'gemini' ? 'GEMINI_API_KEY' : 'OPENAI_API_KEY';
  const key = PropertiesService.getScriptProperties().getProperty(propName);
  if (!key) return { ok: false, error: propName + ' is not set in Script properties.' };
  if (!word) return { ok: false, error: 'Enter a Swedish word first.' };

  const system =
    'You are a Swedish language tutor helping an English-speaking adult learner build a vocabulary list. ' +
    'Given a Swedish word or short phrase, return ONLY a JSON object with exactly these keys. ' +
    'Be concise everywhere — this is a quick reference card, not an essay. Every "example_sv"/"example_en" ' +
    'and both top-level example sentences must be SHORT, natural, everyday sentences of 8 words or fewer. ' +
    'Use the kind of phrasing a native speaker would actually say — prefer common idiomatic expressions over ' +
    'literal-but-unnatural constructions (e.g. say "helt slut" or "jättetrött", not a stiff literal "helt trött", ' +
    'for "very tired"). If in doubt, choose the simpler, more common phrasing an SFI (Swedish for immigrants) course ' +
    'would teach at A2-B1 level.\n' +
    '"Dictionary Form": the word\'s standard dictionary/lemma form, lowercase — infinitive for verbs, singular ' +
    'indefinite for nouns (no article), base/positive form for adjectives; if the input is already the dictionary ' +
    'form, return it unchanged (lowercased); if it\'s inflected (e.g. "gick", "bilar", "snabbare"), return its ' +
    'lemma (e.g. "gå", "bil", "snabb");\n' +
    '"Pronunciation Guide": an English-friendly respelling in lowercase, syllables separated by hyphens, with ONLY the stressed syllable in CAPS, e.g. "hem-TREV-lig";\n' +
    '"Part of Speech": one of ' + JSON.stringify(POS_OPTIONS) + ' (use "noun (en)" or "noun (ett)" to show the gender; pick the most common use);\n' +
    '"English Meaning": a short translation, 6 words or fewer, ALWAYS lowercase (except proper nouns), and for ' +
    'verbs do NOT prefix with "to" (write "take", not "to take") (most common sense; a second sense only if truly common, comma-separated);\n' +
    '"Example Sentence (Swedish)": one short everyday sentence (max 8 words) using the word exactly as given (A2-B1 level);\n' +
    '"Example Sentence (English)": its English translation;\n' +
    '"Forms": an array of the word\'s key grammatical forms, each item shaped ' +
    '{"label": string, "sv": string, "example_sv": string, "example_en": string}, built as follows:\n' +
    '  - If Part of Speech is a noun: exactly 4 items labelled "Singular indefinite", "Singular definite", "Plural indefinite", "Plural definite" ' +
    '(e.g. "en bil", "bilen", "bilar", "bilarna"), each with a short (max 8 word) example sentence using that exact form;\n' +
    '  - If Part of Speech is verb: exactly 4 items labelled "Infinitive", "Present", "Past", "Supine" (the four principal parts, e.g. "välja", "väljer", "valde", "valt"), ' +
    'each with a short (max 8 word) example sentence using that exact form (supine example uses "har" or "hade");\n' +
    '  - If Part of Speech is adjective: exactly 3 items labelled "Base form (en-word)", "Neuter form (ett-word)", "Plural / definite form", each with a short (max 8 word) example sentence;\n' +
    '  - Otherwise (adverb, pronoun, preposition, etc.): an empty array [].\n' +
    'The "sv" field of each form item must be ONLY the inflected word itself (no article, no extra words) so it can be matched against later, e.g. "bilar" not "en bilar". ' +
    'If the input is not a real Swedish word, still return the JSON with "English Meaning" set to "Not recognised as Swedish" and "Forms" as [].';

  const userMessage = sourceSentence
    ? word + '\n\nThe learner found this word in the following sentence — if natural, base "Example Sentence (Swedish)" ' +
      'closely on it (adapt only as needed to fit the length/word rules), and still give its English translation as usual:\n"' + sourceSentence + '"'
    : word;

  let fields;
  if (PROVIDER === 'gemini') {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: 'user', parts: [{ text: userMessage }] }],
        generationConfig: { temperature: 0.3, responseMimeType: 'application/json', maxOutputTokens: AUTOFILL_MAX_TOKENS },
      }),
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) {
      return { ok: false, error: 'Gemini error ' + res.getResponseCode() + ': ' + ((body.error && body.error.message) || 'unknown') };
    }
    fields = JSON.parse(body.candidates[0].content.parts[0].text);
  } else {
    const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        model: OPENAI_MODEL,
        // Newer OpenAI models (the gpt-5.x family, including gpt-5.4-mini) reject
        // the old `max_tokens` param — it must be `max_completion_tokens` now —
        // and only accept the default temperature (1), so `temperature` is
        // omitted here rather than set to 0.3 as it used to be.
        max_completion_tokens: AUTOFILL_MAX_TOKENS,
        response_format: { type: 'json_object' },
        messages: [{ role: 'system', content: system }, { role: 'user', content: userMessage }],
      }),
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) {
      return { ok: false, error: 'OpenAI error ' + res.getResponseCode() + ': ' + ((body.error && body.error.message) || 'unknown') };
    }
    fields = JSON.parse(body.choices[0].message.content);
  }
  // Forms travels to the client as a JSON string, same shape it's stored in the sheet.
  if (fields && Array.isArray(fields.Forms)) {
    fields['Grammar Forms'] = JSON.stringify(fields.Forms);
  } else {
    fields['Grammar Forms'] = '[]';
  }
  delete fields.Forms;
  return { ok: true, fields: fields };
}

/**
 * Writes a short (100-150 word) Swedish reading passage that naturally uses
 * as many of `words` as it can, for the "Today's reading" feature — a pure
 * LLM call, same shape as autofill_ above, with no Sheet read or write.
 */
function generateReading_(words) {
  const propName = PROVIDER === 'gemini' ? 'GEMINI_API_KEY' : 'OPENAI_API_KEY';
  const key = PropertiesService.getScriptProperties().getProperty(propName);
  if (!key) return { ok: false, error: propName + ' is not set in Script properties.' };
  const list = (Array.isArray(words) ? words : []).map(w => String(w || '').trim()).filter(Boolean);
  if (list.length < 4) return { ok: false, error: 'Need at least 4 words to generate a reading passage.' };

  const system =
    'You are a Swedish language tutor writing short graded-reader passages for an English-speaking adult ' +
    'learner at A2-B1 level. Write ONE short passage in Swedish, ' + READING_WORD_COUNT.min + '-' + READING_WORD_COUNT.max +
    ' words long, that naturally uses as many as possible of a given list of Swedish words the learner is ' +
    'currently reviewing. You may inflect/conjugate the given words as needed to fit the sentence (e.g. "springa" ' +
    'may appear as "sprang" or "springer") — natural grammar always wins over using a word\'s exact dictionary form. ' +
    'It is fine to skip a word if it cannot be worked in naturally. Keep the vocabulary and sentence structure ' +
    'otherwise simple and mostly-familiar (short sentences, everyday topics, present or simple past tense) — the ' +
    'point is comfortable reading practice, not a challenge. Return ONLY a JSON object with exactly these keys: ' +
    '"text": the Swedish passage as a single string (use "\\n\\n" between paragraphs if you use more than one); ' +
    '"wordsUsed": a JSON array of the words from the given list you actually managed to include.';

  const userMessage = 'Words to review:\n' + list.join(', ');

  let fields;
  if (PROVIDER === 'gemini') {
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: 'user', parts: [{ text: userMessage }] }],
        generationConfig: { temperature: 0.7, responseMimeType: 'application/json', maxOutputTokens: READING_MAX_TOKENS },
      }),
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) {
      return { ok: false, error: 'Gemini error ' + res.getResponseCode() + ': ' + ((body.error && body.error.message) || 'unknown') };
    }
    fields = JSON.parse(body.candidates[0].content.parts[0].text);
  } else {
    const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        model: OPENAI_MODEL,
        // Same gpt-5.x constraint as autofill_ above: no `temperature`, and
        // `max_completion_tokens` rather than the old `max_tokens`.
        max_completion_tokens: READING_MAX_TOKENS,
        response_format: { type: 'json_object' },
        messages: [{ role: 'system', content: system }, { role: 'user', content: userMessage }],
      }),
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) {
      return { ok: false, error: 'OpenAI error ' + res.getResponseCode() + ': ' + ((body.error && body.error.message) || 'unknown') };
    }
    fields = JSON.parse(body.choices[0].message.content);
  }
  return { ok: true, text: String(fields.text || ''), wordsUsed: Array.isArray(fields.wordsUsed) ? fields.wordsUsed : [] };
}

/**
 * Returns the sheet, creating it and its header row if needed, and appends
 * (at the end of the existing header row) any column from HEADERS that an
 * older sheet is missing. Never reorders existing columns.
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    return sh;
  }
  const lastCol = sh.getLastColumn();
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = HEADERS.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

// Maps header name -> 1-based column index, using the sheet's actual header row.
function headerMap_(sh) {
  const lastCol = sh.getLastColumn();
  const row = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  row.forEach((h, i) => { if (h) map[String(h)] = i + 1; });
  return map;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Reads every row from the sheet. Backed by a short-lived cache (see
// SEARCH_CACHE_SECONDS) because the Sheets read is what makes search feel
// slow — a full sheet scan happens at most once every ~50s per cache slot
// instead of on every keystroke. invalidateCache_() clears it on any write.
function readAll_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('rows_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through to re-read */ }
  }
  const sh = getSheet_();
  const map = headerMap_(sh);
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  const wordCol = map['Swedish Word'] - 1;
  const rows = values.slice(1)
    .filter(r => String(r[wordCol]).trim() !== '')
    .map(r => {
      const o = {};
      HEADERS.forEach(h => { const c = map[h]; o[h] = c ? String(r[c - 1] == null ? '' : r[c - 1]) : ''; });
      return o;
    });
  try { cache.put('rows_v1', JSON.stringify(rows), SEARCH_CACHE_SECONDS); } catch (e) { /* row set too large to cache — fine, just slower */ }
  return rows;
}

function invalidateCache_() {
  try { CacheService.getScriptCache().remove('rows_v1'); } catch (e) {}
}

function findRow_(sh, word) {
  const target = String(word).trim().toLowerCase();
  if (!target) return -1;
  const map = headerMap_(sh);
  const wordCol = map['Swedish Word'];
  const words = sh.getRange(1, wordCol, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < words.length; i++) {
    if (String(words[i][0]).trim().toLowerCase() === target) return i + 1;
  }
  return -1;
}

function norm_(s) {
  return String(s == null ? '' : s).toLowerCase().trim();
}

// Safe JSON-array parse, used for the "Seen As" cell — never throws, always
// returns an array (empty on missing/malformed input).
function safeParseArr_(s) {
  try { const a = JSON.parse(s || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}

// Best-effort Swedish suffix stripper, used only as a fallback for older rows
// that have no "Grammar Forms" saved yet. Longer/more specific suffixes are
// tried first. Deliberately has NO single-letter suffixes (no bare "n", "t",
// "s"): those are too common as the last letter of an unrelated word and
// cause false matches — e.g. "något" (something) and "någon" (any/someone)
// both end in one letter and would otherwise collapse to the same fake stem.
// Works well for the common case (flygplats/flygplatser/flygplatserna,
// bil/bilar/bilarna, ...) and both sides of a comparison are always stemmed
// the same way, so it's self-consistent. The remaining stem must be at least
// 4 characters, trading a little recall for far fewer false hits.
const STEM_SUFFIXES = ['arna', 'orna', 'erna', 'ana', 'ens', 'ets', 'are', 'ast', 'ade', 'and', 'else',
  'or', 'ar', 'er', 'en', 'et', 'na'];
function stem_(word) {
  const w = norm_(word);
  for (let i = 0; i < STEM_SUFFIXES.length; i++) {
    const suf = STEM_SUFFIXES[i];
    if (w.length - suf.length >= 4 && w.slice(-suf.length) === suf) return w.slice(0, -suf.length);
  }
  return w;
}

// Dictionary-style match: a query matches a row if it's found in the saved
// word or meaning, OR it exactly matches (or is contained in) one of the
// word's saved inflected forms (Grammar Forms), OR — as a last-resort
// fallback for rows without saved forms — its stem matches the word's stem.
// This is what lets searching "flygplatserna" find the row saved as "flygplats".
function rowMatchesQuery_(row, q, qStem) {
  if (!q) return true;
  const word = norm_(row['Swedish Word']);
  const meaning = norm_(row['English Meaning']);
  if (word.indexOf(q) !== -1 || meaning.indexOf(q) !== -1) return true;
  try {
    const forms = JSON.parse(row['Grammar Forms'] || '[]');
    for (let i = 0; i < forms.length; i++) {
      const f = norm_(forms[i] && forms[i].sv);
      if (f && (f === q || f.indexOf(q) !== -1)) return true;
    }
  } catch (e) { /* malformed Grammar Forms — ignore */ }
  // "Seen As": the inflected forms the user actually typed before the word got
  // saved under its dictionary form, e.g. searching "gick" finds a row saved as "gå".
  try {
    const seen = JSON.parse(row['Seen As'] || '[]');
    for (let i = 0; i < seen.length; i++) {
      const s = norm_(seen[i]);
      if (s && (s === q || s.indexOf(q) !== -1)) return true;
    }
  } catch (e) { /* malformed Seen As — ignore */ }
  if (qStem && qStem.length >= 4 && stem_(word) === qStem) return true;
  return false;
}

// GET  ?meta=1   -> { ok, provider, model } for the currently configured auto-fill model
// GET  ?q=hund    -> search (Swedish word, English meaning, or any saved inflected form).
//                    Matches "flygplatserna" against a row saved as "flygplats". No q -> all rows.
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.meta) {
      return json_({ ok: true, provider: PROVIDER, model: PROVIDER === 'gemini' ? GEMINI_MODEL : OPENAI_MODEL });
    }
    const q = norm_(p.q || '');
    let rows = readAll_();
    if (q) {
      const qStem = stem_(q);
      rows = rows.filter(r => rowMatchesQuery_(r, q, qStem));
    }
    return json_({ ok: true, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// POST (JSON sent as text/plain). Actions:
//   { action: 'autofill', word, sourceSentence } -> LLM fills the other fields + Grammar Forms
//   { action: 'generateReading', words }         -> LLM writes a short passage using the given due words
//   { action: 'setLearned', word, learned }      -> tick / untick the Learned column
//   { action: 'updateReview', word, reviewState } -> overwrite the Review State column (spaced repetition)
//   { 'Swedish Word': ..., ... }                 -> add a word, or update it if it exists
//                                                    ('Swedish Word' is lowercased on write; 'Seen As' is merged, not overwritten)
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const data = JSON.parse(e.postData.contents);

    const requiredPin = PropertiesService.getScriptProperties().getProperty('APP_PIN');
    if (requiredPin) {
      const given = String(data.pin || '').trim();
      if (!given) return json_({ ok: false, error: 'PIN required' });
      if (given !== requiredPin) return json_({ ok: false, error: 'Incorrect PIN' });
    }

    if (data.action === 'autofill') {
      return json_(autofill_(String(data.word || '').trim(), String(data.sourceSentence || '').trim()));
    }
    if (data.action === 'generateReading') {
      return json_(generateReading_(Array.isArray(data.words) ? data.words : []));
    }

    lock.waitLock(10000);
    const sh = getSheet_();
    const map = headerMap_(sh);

    if (data.action === 'setLearned') {
      const row = findRow_(sh, data.word);
      if (row < 0) return json_({ ok: false, error: 'Word not found in the sheet' });
      sh.getRange(row, map['Learned']).setValue(data.learned ? 'Yes' : '');
      invalidateCache_();
      return json_({ ok: true });
    }

    if (data.action === 'updateReview') {
      const row = findRow_(sh, data.word);
      if (row < 0) return json_({ ok: false, error: 'Word not found in the sheet' });
      sh.getRange(row, map['Review State']).setValue(String(data.reviewState || '{}'));
      invalidateCache_();
      return json_({ ok: true });
    }

    const word = String(data['Swedish Word'] || '').trim();
    if (!word) return json_({ ok: false, error: 'Swedish Word is required' });

    const existingRow = findRow_(sh, word);
    const lastCol = sh.getLastColumn();
    const rowValues = existingRow > 0
      ? sh.getRange(existingRow, 1, 1, lastCol).getValues()[0]
      : new Array(lastCol).fill('');

    HEADERS.forEach(h => {
      const c = map[h]; if (!c) return;
      if (h === 'Learned') {
        if (data.Learned !== undefined) rowValues[c - 1] = String(data.Learned);
        // else: keep whatever was already there (existing value, or '' for a new row)
      } else if (h === 'Swedish Word' && data[h] !== undefined) {
        // Lowercased on write so the stored headword is always the canonical
        // dictionary-form key (matching stays case-insensitive throughout).
        rowValues[c - 1] = String(data[h]).trim().toLowerCase();
      } else if (h === 'Seen As' && data[h] !== undefined) {
        // Merge, don't overwrite — accumulates every inflected form the word
        // has ever been looked up as, e.g. ["gick"] then later ["gått"].
        const merged = Array.from(new Set(safeParseArr_(rowValues[c - 1]).concat(safeParseArr_(data[h]))));
        rowValues[c - 1] = JSON.stringify(merged);
      } else if (data[h] !== undefined) {
        rowValues[c - 1] = String(data[h]).trim();
      }
    });

    if (existingRow > 0) {
      sh.getRange(existingRow, 1, 1, lastCol).setValues([rowValues]);
      invalidateCache_();
      return json_({ ok: true, action: 'updated' });
    }
    sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([rowValues]);
    invalidateCache_();
    return json_({ ok: true, action: 'added' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}