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

/**
 * Passages tab: saved AI-generated reading passages (story + target words +
 * quiz) for the Read tab. A second, independent sheet tab from Words — a
 * passage references many words and a word appears in many passages, so this
 * is a different entity, not more Words columns. Same create-on-first-use /
 * append-missing-header-by-name migration pattern as getSheet_() below
 * (see getPassagesSheet_), reusing headerMap_() unchanged since it's already
 * generic over whatever sheet is passed to it.
 */
const PASSAGES_SHEET_NAME = 'Passages';
const PASSAGES_HEADERS = [
  'Id',               // ISO timestamp string — same id scheme the client already used for reading-history entries
  'Created',          // ISO timestamp, kept as its own column so opening the Sheet doesn't require parsing Id
  'Status',           // 'ready' | 'archived' — no client-only transient state (loading/error) is ever persisted here
  'Title',
  'Level',
  'WordCount',
  'Paragraphs',       // JSON array of strings, one per paragraph
  'Targets',          // JSON array of {wordId, lemma, surface, paragraphIndex, pos} — wordId is just the lemma;
                       // there's no separate word-id concept anywhere in this app, rows are keyed by headword
  'Quiz',             // JSON array of {prompt, options, answerIndex, wordId}
  'RequestedWords',   // JSON array of the headwords fed into the generation prompt
  'OpenedWords',      // JSON array, append-only/de-duped — words the user has tapped open in the reading side panel
  'QuizResult',       // JSON {score, total, answers, completedAt}, or '' if the quiz hasn't been taken yet
  'Source',           // 'auto' | 'manual-regenerate' | 'migrated'
];
// Separate from SEARCH_CACHE_SECONDS' 'rows_v1' key so Words and Passages
// caching never interfere with each other.
const PASSAGES_CACHE_KEY = 'passages_v1';

// ---- Auto-fill settings: check these two match what you use ----
const PROVIDER = 'openai';                       // 'openai' or 'gemini'
const OPENAI_MODEL = 'gpt-5.4-mini';
const GEMINI_MODEL = 'gemini-3.5-flash-lite';
const POS_OPTIONS = ['noun (en)', 'noun (ett)', 'verb', 'adjective', 'adverb', 'pronoun',
  'preposition', 'conjunction', 'interjection', 'numeral', 'phrase'];

// Hard cap on generated tokens. Keeps every auto-fill call short (fields are
// meant to be quick reference notes, not essays) and cuts response time a lot,
// since latency scales mostly with how many tokens the model has to generate.
const AUTOFILL_MAX_TOKENS = 550;   // room for 5 adjective forms + Recognised/Suggestion/Verdict keys

// Same rationale as AUTOFILL_MAX_TOKENS, sized for a 100-150 word Swedish
// passage plus the wordsUsed array and JSON overhead.
const READING_MAX_TOKENS = 700;
const READING_WORD_COUNT = { min: 100, max: 150 };

// generatePassage_ produces a much bigger payload than either call above —
// a full story PLUS a per-word targets array PLUS a quiz — so it gets its
// own, larger token budget and its own word-count/quiz-size ranges.
const PASSAGE_MAX_TOKENS = 2200;
const PASSAGE_WORD_COUNT = { min: 120, max: 220 };
const PASSAGE_QUIZ_COUNT = { min: 4, max: 6 };

// How long a search result set is cached (seconds). Search re-reads the whole
// sheet otherwise, and a Sheets API read is the slowest part of a search by
// far, so this is the main lever for making search feel instant. Any
// save/edit/learn immediately clears the cache so you never see stale data.
const SEARCH_CACHE_SECONDS = 50;

// savedEntry (optional) switches on "check mode", used by the client's
// "Check my words" tool: the model also judges the saved entry and returns
// "Verdict" ("ok" | "fix") and "Issues" alongside a corrected entry.
function autofill_(word, sourceSentence, savedEntry) {
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
    'IMPORTANT: the entry is ALWAYS about the Dictionary Form, never the form that was typed. If the input is ' +
    'inflected (e.g. "tystare"), every other field — pronunciation, part of speech, meaning, example and forms — ' +
    'describes the dictionary form ("tyst" = "quiet", not "quieter").\n' +
    '"Recognised": true if the input is a real Swedish word or phrase (any inflection, correctly spelled), else false;\n' +
    '"Suggestion": if "Recognised" is false, the most likely Swedish word the learner meant (e.g. "flygplats" for ' +
    '"flyplats"), in its dictionary form; otherwise "";\n' +
    '"Dictionary Form": the word\'s standard dictionary/lemma form, lowercase — infinitive for verbs, singular ' +
    'indefinite for nouns (no article), base/positive form for adjectives; if the input is already the dictionary ' +
    'form, return it unchanged (lowercased); if it\'s inflected (e.g. "gick", "bilar", "snabbare"), return its ' +
    'lemma (e.g. "gå", "bil", "snabb");\n' +
    '"Pronunciation Guide": an English-friendly respelling of the dictionary form, all lowercase, syllables separated by hyphens, ' +
    'with the stress mark "ˈ" directly before the stressed syllable — never capital letters — e.g. "hem-ˈtrev-lig", "ˈtyst";\n' +
    '"Part of Speech": one of ' + JSON.stringify(POS_OPTIONS) + ' (use "noun (en)" or "noun (ett)" to show the gender; pick the most common use);\n' +
    '"English Meaning": a short translation of the dictionary form, 6 words or fewer, ALWAYS lowercase (except proper nouns); ' +
    'nouns in the singular ("sandwich", not "sandwiches"); verbs in the plain form and NOT prefixed with "to" ' +
    '(write "take", not "to take" or "took") (most common sense; a second sense only if truly common, comma-separated);\n' +
    '"Example Sentence (Swedish)": one short everyday sentence (max 8 words) using the dictionary form (A2-B1 level), ' +
    'unless the learner supplies a source sentence (then follow that, even if it uses another form);\n' +
    '"Example Sentence (English)": its English translation;\n' +
    '"Forms": an array of the word\'s key grammatical forms, each item shaped ' +
    '{"label": string, "sv": string, "example_sv": string, "example_en": string}, built as follows:\n' +
    '  - If Part of Speech is a noun: exactly 4 items labelled "Singular indefinite", "Singular definite", "Plural indefinite", "Plural definite" ' +
    '(e.g. "en bil", "bilen", "bilar", "bilarna"), each with a short (max 8 word) example sentence using that exact form;\n' +
    '  - If Part of Speech is verb: exactly 4 items labelled "Infinitive", "Present", "Past", "Supine" (the four principal parts, e.g. "välja", "väljer", "valde", "valt"), ' +
    'each with a short (max 8 word) example sentence using that exact form (supine example uses "har" or "hade");\n' +
    '  - If Part of Speech is adjective: exactly 5 items labelled "Base form (en-word)", "Neuter form (ett-word)", "Plural / definite form", ' +
    '"Comparative", "Superlative" (e.g. "tyst", "tyst", "tysta", "tystare", "tystast"; irregular ones as they are, e.g. "bättre", "bäst"; ' +
    'for adjectives that only compare with "mer"/"mest", use e.g. "mer intressant"), each with a short (max 8 word) example sentence;\n' +
    '  - Otherwise (adverb, pronoun, preposition, etc.): an empty array [].\n' +
    'The "sv" field of each form item must be ONLY the inflected word itself (no article, no extra words) so it can be matched against later, e.g. "bilar" not "en bilar". ' +
    'If the input is not a real Swedish word, still return the JSON with "Recognised" false, "English Meaning" set to "Not recognised as Swedish" and "Forms" as [].' +
    (savedEntry
      ? '\nThe learner also sends an entry they SAVED earlier for this word. Compare it with your own answer and add two keys: ' +
        '"Verdict": "fix" only if the saved entry has a real error — headword is not the dictionary form, a misspelling, ' +
        'the meaning/part of speech is wrong or belongs to a different form, the example uses the word wrongly, or forms are ' +
        'missing or wrong — otherwise "ok" (ignore harmless wording differences, synonyms and pronunciation style); ' +
        '"Issues": one short sentence naming the errors when "fix", else "".'
      : '');

  let userMessage = sourceSentence
    ? word + '\n\nThe learner found this word in the following sentence — if natural, base "Example Sentence (Swedish)" ' +
      'closely on it (adapt only as needed to fit the length/word rules), and still give its English translation as usual:\n"' + sourceSentence + '"'
    : word;
  if (savedEntry) userMessage += '\n\nSaved entry:\n' + JSON.stringify(savedEntry);

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
 * Writes a longer Swedish reading passage (PASSAGE_WORD_COUNT words) grounded
 * in `words`, plus a "targets" array pinning exactly where each word appears
 * (paragraph index + its exact inflected surface form) and a short multiple-
 * choice quiz over it — the one structured call behind the Read tab's
 * "Generate passage" action. Unlike autofill_/generateReading_ above, this
 * uses each provider's real structured-output/schema mode (not just JSON-
 * mode prompting, see callPassageProvider_) AND validates the result
 * server-side afterward (validatePassage_ below) — schema mode guarantees
 * *shape*, not *correctness* (it can't guarantee a target's surface actually
 * occurs where claimed, or that a lemma traces back to a word actually asked
 * for), so both are needed. Retries once — a fresh call, same prompt — if
 * validation still fails after dropping whatever it safely can; no existing
 * call in this file retries today, so this is new territory here.
 */
function generatePassage_(words, level) {
  const propName = PROVIDER === 'gemini' ? 'GEMINI_API_KEY' : 'OPENAI_API_KEY';
  const key = PropertiesService.getScriptProperties().getProperty(propName);
  if (!key) return { ok: false, error: propName + ' is not set in Script properties.' };
  const list = (Array.isArray(words) ? words : []).map(w => String(w || '').trim()).filter(Boolean);
  if (list.length < 4) return { ok: false, error: 'Need at least 4 words to generate a passage.' };
  const lvl = String(level || 'A2').trim() || 'A2';

  for (let attempt = 0; attempt < 2; attempt++) {
    const result = callPassageProvider_(key, list, lvl);
    if (!result.ok) return result;   // hard provider/network error — no point retrying
    const validated = validatePassage_(result.parsed, list);
    if (validated.ok) return { ok: true, data: validated.data };
    if (attempt === 1) return { ok: false, error: 'Generated passage failed validation twice: ' + validated.reason };
  }
}

// The actual provider round trip for generatePassage_ — split out so the
// retry loop above doesn't duplicate the OpenAI/Gemini branching.
function callPassageProvider_(key, list, level) {
  const system =
    'You are a Swedish language tutor writing a short graded-reader story for an English-speaking adult learner at ' +
    level + ' level. Write ONE passage in Swedish, ' + PASSAGE_WORD_COUNT.min + '-' + PASSAGE_WORD_COUNT.max +
    ' words long, across 2-4 paragraphs, that naturally uses as many as possible of a given list of Swedish words. ' +
    'You may inflect/conjugate a given word as needed to fit the sentence (natural grammar always wins over using ' +
    'a word\'s exact dictionary/lemma form) — it is fine to skip a word if it cannot be worked in naturally. Keep ' +
    'everything outside the target words at or below ' + level + ' level (short sentences, everyday topics, present ' +
    'or simple past tense) — the point is comfortable reading practice, not a challenge.\n' +
    'Then write ' + PASSAGE_QUIZ_COUNT.min + '-' + PASSAGE_QUIZ_COUNT.max + ' multiple-choice questions about the ' +
    'passage: a mix of word-meaning questions and at least one question testing overall understanding of the story. ' +
    'Each question has exactly 4 short answer options with exactly one correct answer.\n' +
    'Return a JSON object with these keys:\n' +
    '"title": a short title for the story, in Swedish;\n' +
    '"paragraphs": an array of the story\'s paragraphs as plain strings (no markdown, no numbering);\n' +
    '"targets": for EVERY word from the given list you actually managed to use, one object ' +
    '{"lemma": the word exactly as given, "surface": the EXACT inflected form you used, character for character, ' +
    'as it appears in "paragraphs", "paragraphIndex": the 0-based index of the paragraph it appears in, ' +
    '"pos": its part of speech (one short word, e.g. "verb", "noun", "adjective", "adverb")} — this is how the ' +
    'reader app highlights the word later, so "surface"/"paragraphIndex" must be exactly right, not approximate;\n' +
    '"quiz": an array of question objects {"prompt": the question, in Swedish or English as fits, "options": ' +
    'exactly 4 short answer strings, "answerIndex": the 0-based index of the correct option, "wordId": the exact ' +
    '"lemma" of the target word this question is about, or "" for a general understanding question}.';

  const userMessage = 'Words to use:\n' + list.join(', ');

  if (PROVIDER === 'gemini') {
    // Gemini's schema dialect: uppercase type names, no additionalProperties.
    const schema = {
      type: 'OBJECT',
      properties: {
        title: { type: 'STRING' },
        paragraphs: { type: 'ARRAY', items: { type: 'STRING' } },
        targets: {
          type: 'ARRAY',
          items: {
            type: 'OBJECT',
            properties: {
              lemma: { type: 'STRING' }, surface: { type: 'STRING' },
              paragraphIndex: { type: 'INTEGER' }, pos: { type: 'STRING' },
            },
            required: ['lemma', 'surface', 'paragraphIndex', 'pos'],
          },
        },
        quiz: {
          type: 'ARRAY',
          items: {
            type: 'OBJECT',
            properties: {
              prompt: { type: 'STRING' }, options: { type: 'ARRAY', items: { type: 'STRING' } },
              answerIndex: { type: 'INTEGER' }, wordId: { type: 'STRING' },
            },
            required: ['prompt', 'options', 'answerIndex', 'wordId'],
          },
        },
      },
      required: ['title', 'paragraphs', 'targets', 'quiz'],
    };
    const res = UrlFetchApp.fetch(
      'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      muteHttpExceptions: true,
      payload: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: 'user', parts: [{ text: userMessage }] }],
        generationConfig: {
          temperature: 0.7, responseMimeType: 'application/json', responseSchema: schema,
          maxOutputTokens: PASSAGE_MAX_TOKENS,
        },
      }),
    });
    const body = JSON.parse(res.getContentText());
    if (res.getResponseCode() !== 200) {
      return { ok: false, error: 'Gemini error ' + res.getResponseCode() + ': ' + ((body.error && body.error.message) || 'unknown') };
    }
    return { ok: true, parsed: JSON.parse(body.candidates[0].content.parts[0].text) };
  }

  // OpenAI strict json_schema mode requires additionalProperties:false and
  // every property listed in required (no optional keys) on every object.
  const schema = {
    type: 'object',
    properties: {
      title: { type: 'string' },
      paragraphs: { type: 'array', items: { type: 'string' } },
      targets: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            lemma: { type: 'string' }, surface: { type: 'string' },
            paragraphIndex: { type: 'integer' }, pos: { type: 'string' },
          },
          required: ['lemma', 'surface', 'paragraphIndex', 'pos'],
          additionalProperties: false,
        },
      },
      quiz: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            prompt: { type: 'string' }, options: { type: 'array', items: { type: 'string' } },
            answerIndex: { type: 'integer' }, wordId: { type: 'string' },
          },
          required: ['prompt', 'options', 'answerIndex', 'wordId'],
          additionalProperties: false,
        },
      },
    },
    required: ['title', 'paragraphs', 'targets', 'quiz'],
    additionalProperties: false,
  };
  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model: OPENAI_MODEL,
      // Same gpt-5.x constraint as autofill_/generateReading_ above: no
      // `temperature`, and `max_completion_tokens` rather than `max_tokens`.
      max_completion_tokens: PASSAGE_MAX_TOKENS,
      response_format: { type: 'json_schema', json_schema: { name: 'passage', strict: true, schema: schema } },
      messages: [{ role: 'system', content: system }, { role: 'user', content: userMessage }],
    }),
  });
  const body = JSON.parse(res.getContentText());
  if (res.getResponseCode() !== 200) {
    return { ok: false, error: 'OpenAI error ' + res.getResponseCode() + ': ' + ((body.error && body.error.message) || 'unknown') };
  }
  return { ok: true, parsed: JSON.parse(body.choices[0].message.content) };
}

/**
 * Server-side semantic checks on a generatePassage_ response — schema mode
 * (above) guarantees shape, not correctness. Drops individual targets/quiz
 * items that fail a check rather than failing the whole passage, except when
 * too little survives to be useful (the two "hard failure" conditions near
 * the end), which is what triggers generatePassage_'s one retry.
 */
function validatePassage_(parsed, requestedWords) {
  const paragraphs = Array.isArray(parsed && parsed.paragraphs) ? parsed.paragraphs.map(p => String(p || '')) : [];
  if (!paragraphs.length) return { ok: false, reason: 'no paragraphs returned' };

  const requestedStems = (requestedWords || []).map(w => stem_(w));

  const targets = (Array.isArray(parsed.targets) ? parsed.targets : []).filter(t => {
    if (!t || typeof t !== 'object') return false;
    const pi = t.paragraphIndex;
    if (!Number.isInteger(pi) || pi < 0 || pi >= paragraphs.length) return false;
    const surface = String(t.surface || '');
    if (!surface) return false;
    // Unicode-aware whole-word check: split the paragraph on runs of
    // non-letters (not ASCII \b, which mis-bounds on å/ä/ö — the same caveat
    // already flagged for the client-side cloze drill) and require an exact
    // case-insensitive match among the paragraph's own tokens, not a
    // substring (indexOf would false-positive on e.g. "sprang" inside a
    // longer word, and false-negative at a å/ä/ö boundary).
    const tokens = paragraphs[pi].match(/[\p{L}]+/gu) || [];
    if (!tokens.some(tok => tok.toLowerCase() === surface.toLowerCase())) return false;
    // The lemma must trace back to a word actually requested (stem-matched,
    // same fallback rowMatchesQuery_ uses elsewhere) — catches the model
    // inventing a target it wasn't given.
    const lemma = String(t.lemma || '');
    if (!lemma) return false;
    const lemmaOk = requestedWords.some(w => norm_(w) === norm_(lemma)) || requestedStems.indexOf(stem_(lemma)) !== -1;
    return lemmaOk;
  }).map(t => ({
    wordId: norm_(t.lemma), lemma: norm_(t.lemma), surface: String(t.surface),
    paragraphIndex: t.paragraphIndex, pos: String(t.pos || ''),
  }));

  const quiz = (Array.isArray(parsed.quiz) ? parsed.quiz : []).filter(q => {
    if (!q || typeof q !== 'object') return false;
    if (!String(q.prompt || '').trim()) return false;
    if (!Array.isArray(q.options) || q.options.length < 2) return false;
    if (!Number.isInteger(q.answerIndex) || q.answerIndex < 0 || q.answerIndex >= q.options.length) return false;
    return true;
  }).map(q => ({
    prompt: String(q.prompt), options: q.options.map(o => String(o)),
    answerIndex: q.answerIndex, wordId: norm_(q.wordId || ''),
  }));

  if (targets.length < 2) return { ok: false, reason: 'fewer than 2 usable target words survived validation' };
  if (quiz.length < PASSAGE_QUIZ_COUNT.min) {
    return { ok: false, reason: 'fewer than ' + PASSAGE_QUIZ_COUNT.min + ' usable quiz questions survived validation' };
  }

  return {
    ok: true,
    data: { title: String(parsed.title || '').trim() || 'Reading passage', paragraphs: paragraphs, targets: targets, quiz: quiz },
  };
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

// Same as getSheet_() above, but for the Passages tab — identical create-on-
// first-use / append-missing-header-by-name logic, just a different sheet
// name and header list. Kept as its own function rather than parameterizing
// getSheet_() itself, so a change to one never risks the other.
function getPassagesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PASSAGES_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(PASSAGES_SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(PASSAGES_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, PASSAGES_HEADERS.length).setFontWeight('bold');
    return sh;
  }
  const lastCol = sh.getLastColumn();
  const existing = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = PASSAGES_HEADERS.filter(h => existing.indexOf(h) === -1);
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

// Same caching shape as readAll_()/invalidateCache_() above, but for the
// Passages tab, under its own cache key (PASSAGES_CACHE_KEY) — a deliberate
// parallel function rather than a generalized/parameterized readAll_(), so a
// change to one read path can never accidentally affect the other's cache key.
function readAllPassages_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(PASSAGES_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through to re-read */ }
  }
  const sh = getPassagesSheet_();
  const map = headerMap_(sh);
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(1, 1, sh.getLastRow(), lastCol).getValues();
  const idCol = map['Id'] - 1;
  const rows = values.slice(1)
    .filter(r => String(r[idCol]).trim() !== '')
    .map(r => {
      const o = {};
      PASSAGES_HEADERS.forEach(h => { const c = map[h]; o[h] = c ? String(r[c - 1] == null ? '' : r[c - 1]) : ''; });
      return o;
    });
  try { cache.put(PASSAGES_CACHE_KEY, JSON.stringify(rows), SEARCH_CACHE_SECONDS); } catch (e) { /* row set too large to cache — fine, just slower */ }
  return rows;
}

function invalidatePassagesCache_() {
  try { CacheService.getScriptCache().remove(PASSAGES_CACHE_KEY); } catch (e) {}
}

// Same shape as findRow_() below, but matches the Passages tab's 'Id' column
// instead of 'Swedish Word'. Ids are timestamp strings, not user-typed text,
// so this compares as-is (trimmed) rather than lowercasing.
function findPassageRow_(sh, id) {
  const target = String(id == null ? '' : id).trim();
  if (!target) return -1;
  const map = headerMap_(sh);
  const idCol = map['Id'];
  if (!idCol) return -1;
  const ids = sh.getRange(1, idCol, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === target) return i + 1;
  }
  return -1;
}

// Appends a full Passage object (keyed by PASSAGES_HEADERS names, with any
// JSON array/object fields already JSON.stringify'd by the caller) as a new
// row. Shared by the one-time client migration ('importPassage' below) and,
// later, passage generation — both just need "write a fully-formed passage
// row", never a partial one, so one append helper covers both.
function appendPassageRow_(sh, map, passage) {
  const lastCol = sh.getLastColumn();
  const rowValues = new Array(lastCol).fill('');
  PASSAGES_HEADERS.forEach(h => {
    const c = map[h]; if (!c) return;
    const v = passage[h];
    rowValues[c - 1] = v === undefined || v === null ? '' : String(v);
  });
  sh.getRange(sh.getLastRow() + 1, 1, 1, lastCol).setValues([rowValues]);
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

// Folds sourceRow's Learned/Review State/Seen As into targetRow before the
// source row is deleted (true-duplicate merge only, in doPost's upsert).
// Seen As is unioned; Learned/Review State keep target's value if it already
// has one, else fall back to source's.
function mergeRowInto_(sh, map, targetRow, sourceRow) {
  ['Learned', 'Review State', 'Seen As'].forEach(h => {
    const c = map[h]; if (!c) return;
    const targetCell = sh.getRange(targetRow, c);
    const sourceVal = sh.getRange(sourceRow, c).getValue();
    if (h === 'Seen As') {
      const merged = Array.from(new Set(safeParseArr_(targetCell.getValue()).concat(safeParseArr_(sourceVal))));
      targetCell.setValue(JSON.stringify(merged));
    } else if (!String(targetCell.getValue()).trim()) {
      targetCell.setValue(sourceVal);
    }
  });
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

// GET  ?meta=1       -> { ok, provider, model } for the currently configured auto-fill model
// GET  ?passages=1   -> { ok, passages } — light list (no Paragraphs/Targets/Quiz) for the Read library screen
// GET  ?passage=<id> -> { ok, passage } — one full passage, incl. Paragraphs/Targets/Quiz, for reading/quiz screens
// GET  ?q=hund       -> search (Swedish word, English meaning, or any saved inflected form).
//                       Matches "flygplatserna" against a row saved as "flygplats". No q -> all rows.
function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.meta) {
      return json_({ ok: true, provider: PROVIDER, model: PROVIDER === 'gemini' ? GEMINI_MODEL : OPENAI_MODEL });
    }
    if (p.passages) {
      const list = readAllPassages_().map(function (row) {
        return {
          Id: row.Id, Created: row.Created, Status: row.Status, Title: row.Title,
          Level: row.Level, WordCount: row.WordCount,
          RequestedCount: safeParseArr_(row.RequestedWords).length,
          OpenedCount: safeParseArr_(row.OpenedWords).length,
          QuizResult: row.QuizResult,
        };
      });
      return json_({ ok: true, passages: list });
    }
    if (p.passage) {
      const target = String(p.passage).trim();
      const found = readAllPassages_().find(function (row) { return row.Id === target; });
      if (!found) return json_({ ok: false, error: 'Passage not found' });
      return json_({ ok: true, passage: found });
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
//   { action: 'autofill', word, sourceSentence, savedEntry? } -> LLM fills the other fields + Grammar Forms
//                                                    (savedEntry = check mode: also returns Verdict/Issues)
//   { action: 'delete', word }                   -> remove the word's row
//   { action: 'generateReading', words }         -> LLM writes a short passage using the given due words
//   { action: 'generatePassage', words, level? } -> LLM writes a validated story+targets+quiz passage, saved to
//                                                    the Passages tab; { ok:false, error } if generation/validation fails
//   { action: 'importPassage', passage }         -> append an already-built Passage row (one-time client migration
//                                                    of the old local reading history; no-ops if its Id already exists)
//   { action: 'updatePassage', id, patch }       -> patch a passage's openedWords/quizResult/status only
//                                                    (never its Paragraphs/Targets/Quiz)
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
      const savedEntry = data.savedEntry && typeof data.savedEntry === 'object' ? data.savedEntry : null;
      return json_(autofill_(String(data.word || '').trim(), String(data.sourceSentence || '').trim(), savedEntry));
    }
    if (data.action === 'generateReading') {
      return json_(generateReading_(Array.isArray(data.words) ? data.words : []));
    }

    // Unlike the two branches above, this one writes to the Sheet (the
    // Passages tab, not Words), so it needs the lock — but only around its
    // own append step, never the Words-sheet lock/getSheet_() call below.
    if (data.action === 'importPassage') {
      const passage = data.passage && typeof data.passage === 'object' ? data.passage : null;
      if (!passage || !String(passage.Id || '').trim()) return json_({ ok: false, error: 'A passage with an Id is required' });
      lock.waitLock(10000);
      const psh = getPassagesSheet_();
      const pmap = headerMap_(psh);
      if (findPassageRow_(psh, passage.Id) > 0) {
        // Already imported (e.g. a retried migration after a partial earlier
        // failure) — treat as success, not an error, so retries are safe.
        return json_({ ok: true, passage: passage, action: 'skipped' });
      }
      appendPassageRow_(psh, pmap, passage);
      invalidatePassagesCache_();
      return json_({ ok: true, passage: passage, action: 'imported' });
    }

    // generatePassage_ itself is an unlocked LLM call (same rationale as
    // autofill/generateReading above), but persisting its result is a Sheet
    // write, so — same pattern as importPassage just above — this action
    // acquires the lock itself only for the append step, and returns before
    // ever reaching the Words-sheet lock/getSheet_() call below.
    if (data.action === 'generatePassage') {
      const words = Array.isArray(data.words) ? data.words : [];
      const result = generatePassage_(words, data.level);
      if (!result.ok) return json_(result);
      const now = new Date().toISOString();
      const passage = {
        Id: now,
        Created: now,
        Status: 'ready',
        Title: result.data.title,
        Level: String(data.level || 'A2').trim() || 'A2',
        WordCount: String(result.data.paragraphs.join(' ').split(/\s+/).filter(Boolean).length),
        Paragraphs: JSON.stringify(result.data.paragraphs),
        Targets: JSON.stringify(result.data.targets),
        Quiz: JSON.stringify(result.data.quiz),
        RequestedWords: JSON.stringify(words),
        OpenedWords: '[]',
        QuizResult: '',
        Source: data.force ? 'manual-regenerate' : 'auto',
      };
      lock.waitLock(10000);
      const psh = getPassagesSheet_();
      const pmap = headerMap_(psh);
      appendPassageRow_(psh, pmap, passage);
      invalidatePassagesCache_();
      return json_({ ok: true, passage: passage });
    }

    // Same lock-just-for-this-step pattern as importPassage/generatePassage
    // above — targeted single-row patch on the Passages sheet, mirroring how
    // setLearned/updateReview below do a targeted single-cell write on Words,
    // except this touches the row's own sheet (Passages, not Words) so it
    // can't reuse the sh/map the locked block below sets up.
    if (data.action === 'updatePassage') {
      const id = String(data.id || '').trim();
      const patch = data.patch && typeof data.patch === 'object' ? data.patch : {};
      if (!id) return json_({ ok: false, error: 'A passage id is required' });
      lock.waitLock(10000);
      const psh = getPassagesSheet_();
      const pmap = headerMap_(psh);
      const row = findPassageRow_(psh, id);
      if (row < 0) return json_({ ok: false, error: 'Passage not found' });
      // Only ever patches these three fields — Paragraphs/Targets/Quiz (the
      // generated content itself) are never touched by this action.
      if (patch.openedWords !== undefined && pmap['OpenedWords']) {
        psh.getRange(row, pmap['OpenedWords']).setValue(JSON.stringify(Array.isArray(patch.openedWords) ? patch.openedWords : []));
      }
      if (patch.quizResult !== undefined && pmap['QuizResult']) {
        psh.getRange(row, pmap['QuizResult']).setValue(patch.quizResult ? JSON.stringify(patch.quizResult) : '');
      }
      if (patch.status !== undefined && pmap['Status']) {
        psh.getRange(row, pmap['Status']).setValue(String(patch.status || ''));
      }
      invalidatePassagesCache_();
      return json_({ ok: true });
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

    if (data.action === 'delete') {
      const row = findRow_(sh, data.word);
      if (row < 0) return json_({ ok: false, error: 'Word not found in the sheet' });
      sh.deleteRow(row);
      invalidateCache_();
      return json_({ ok: true });
    }

    const word = String(data['Swedish Word'] || '').trim();
    if (!word) return json_({ ok: false, error: 'Swedish Word is required' });
    // Control field only — the word this row used to be saved as, when a
    // regenerate/edit changed the headword. Never one of HEADERS, so the
    // HEADERS.forEach overlay below can never write it into a cell.
    const previousWord = String(data.previousWord || '').trim();
    const lastCol = sh.getLastColumn();

    // (a) a row exists for the NEW word -> update it, merging in and deleting
    // any different row previousWord also matches (true-duplicate cleanup).
    // (b) no row for the new word, but previousWord matches an existing row ->
    // rename/overwrite that row in place (the bug-fix path: e.g. regenerating
    // "saknade" into its dictionary form "sakna" must not orphan the old row).
    // (c) neither -> insert a new row, same as today.
    let targetRow = findRow_(sh, word);

    if (targetRow > 0) {
      if (previousWord && norm_(previousWord) !== norm_(word)) {
        const priorRow = findRow_(sh, previousWord);
        if (priorRow > 0 && priorRow !== targetRow) {
          mergeRowInto_(sh, map, targetRow, priorRow);
          sh.deleteRow(priorRow);
          if (priorRow < targetRow) targetRow -= 1;   // rows below a deleted row shift up
        }
      }
    } else if (previousWord) {
      const priorRow = findRow_(sh, previousWord);
      if (priorRow > 0) targetRow = priorRow;
    }

    const rowValues = targetRow > 0
      ? sh.getRange(targetRow, 1, 1, lastCol).getValues()[0]
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

    if (targetRow > 0) {
      sh.getRange(targetRow, 1, 1, lastCol).setValues([rowValues]);
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