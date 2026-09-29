#!/usr/bin/env node
// Pre-ship checks for ordbok. No dependencies. Prints one PASS/WARN/FAIL line
// per check and exits 1 if anything FAILed. Run from anywhere in the repo:
//   node .claude/skills/ship/check.mjs
import { readFileSync, writeFileSync, mkdtempSync, rmSync } from 'node:fs';
import { execFileSync, spawnSync } from 'node:child_process';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const root = execFileSync('git', ['rev-parse', '--show-toplevel'], { encoding: 'utf8' }).trim();
const read = f => readFileSync(join(root, f), 'utf8');
const html = read('index.html');
const gs = read('code.gs');
const tmp = mkdtempSync(join(tmpdir(), 'ordbok-ship-'));

let failed = 0;
const report = (status, name, detail = []) => {
  if (status === 'FAIL') failed++;
  console.log(`${status.padEnd(4)}  ${name}`);
  detail.slice(0, 15).forEach(d => console.log(`        ${d}`));
  if (detail.length > 15) console.log(`        …and ${detail.length - 15} more`);
};
const lineOf = (text, idx) => text.slice(0, idx).split('\n').length;

// Inline <script> blocks, with everything else blanked out line-for-line so
// `node --check` line numbers are index.html line numbers.
const scripts = [];
const scriptRe = /<script(\s[^>]*)?>([\s\S]*?)<\/script>/g;
for (let m; (m = scriptRe.exec(html));) {
  if (m[1] && /\bsrc=/.test(m[1])) continue;
  scripts.push({ start: m.index + m[0].indexOf('>') + 1, body: m[2] });
}
const jsLines = html.split('\n').map(() => '');
scripts.forEach(s => {
  const first = lineOf(html, s.start) - 1;
  s.body.split('\n').forEach((l, i) => { jsLines[first + i] = (i === 0 ? jsLines[first + i] : '') + l; });
});
const js = jsLines.join('\n');
const markup = scripts.reduceRight((h, s) => h.slice(0, s.start) + s.body.replace(/[^\n]/g, ' ') + h.slice(s.start + s.body.length), html);

function syntaxCheck(name, source, file) {
  const p = join(tmp, file);
  writeFileSync(p, source);
  const r = spawnSync(process.execPath, ['--check', p], { encoding: 'utf8' });
  if (r.status === 0) return report('PASS', name);
  // Strip the temp dir (which may appear via a symlinked /private path).
  const shown = file.replace('.check.js', '');
  const lines = (r.stderr || '').split('\n').filter(Boolean).map(l => l.replace(/\S*\.check\.js/g, shown));
  report('FAIL', name, lines.slice(0, 6));
}

// 1-2. Syntax
syntaxCheck('JS syntax: index.html <script> blocks', js, 'index.html.check.js');
syntaxCheck('JS syntax: code.gs', gs, 'code.gs.check.js');

// 3. Every $('id') lookup has a matching id="..." somewhere (markup or a template string).
{
  const ids = new Set([...html.matchAll(/\bid="([\w-]+)"/g)].map(m => m[1]));
  const missing = [...js.matchAll(/\$\('([\w-]+)'\)/g)]
    .filter(m => !ids.has(m[1]))
    .map(m => `$('${m[1]}') at index.html:${lineOf(js, m.index)} — no element with that id`);
  report(missing.length ? 'FAIL' : 'PASS', "DOM ids: every $('id') exists", [...new Set(missing)]);
}

// 4. No duplicate ids in static markup.
{
  const seen = new Map(), dups = [];
  for (const m of markup.matchAll(/\bid="([\w-]+)"/g)) {
    if (seen.has(m[1])) dups.push(`id="${m[1]}" at index.html:${seen.get(m[1])} and :${lineOf(markup, m.index)}`);
    else seen.set(m[1], lineOf(markup, m.index));
  }
  report(dups.length ? 'FAIL' : 'PASS', 'Duplicate ids in markup', dups);
}

// 5. localStorage only touched inside the `store` helper.
{
  const start = js.indexOf('const store = {');
  const end = start === -1 ? -1 : js.indexOf('\n};', start);
  const bad = [...js.matchAll(/\blocalStorage\./g)]
    .filter(m => start === -1 || m.index < start || m.index > end)
    .map(m => `index.html:${lineOf(js, m.index)} — use store.get/set/remove instead`);
  report(bad.length ? 'FAIL' : 'PASS', 'localStorage only via store', bad);
}

// 6. STEM_SUFFIXES identical in index.html and code.gs.
{
  const suffixes = src => {
    const m = src.match(/const STEM_SUFFIXES\s*=\s*\[([\s\S]*?)\]/);
    return m ? [...m[1].matchAll(/['"]([^'"]*)['"]/g)].map(x => x[1]) : null;
  };
  const a = suffixes(html), b = suffixes(gs);
  if (!a || !b) report('FAIL', 'STEM_SUFFIXES in sync', ['could not find STEM_SUFFIXES in ' + (!a ? 'index.html' : 'code.gs')]);
  else if (a.join() !== b.join()) report('FAIL', 'STEM_SUFFIXES in sync', [`index.html: ${a.join(', ')}`, `code.gs:    ${b.join(', ')}`]);
  else report('PASS', 'STEM_SUFFIXES in sync');
}

// Files that would be committed: tracked + untracked-but-not-ignored.
const files = execFileSync('git', ['ls-files', '--cached', '--others', '--exclude-standard'], { cwd: root, encoding: 'utf8' })
  .split('\n').filter(Boolean).filter(f => !/\.(png|jpe?g|gif|ico|webp|woff2?|pdf)$/i.test(f));
const texts = files.map(f => { try { return [f, read(f)]; } catch { return [f, '']; } });

// 7. Secrets. Patterns are assembled from parts so this file doesn't match itself.
{
  const patterns = [
    ['OpenAI-style key', new RegExp('s' + 'k-[A-Za-z0-9_-]{20,}')],
    ['Google API key', new RegExp('AI' + 'za[0-9A-Za-z_-]{30,}')],
    ['hardcoded *_API_KEY', new RegExp('[A-Z_]*_API' + '_KEY\\s*[:=]\\s*[\'"][^\'"]{8,}[\'"]')],
    ['hardcoded APP_PIN', new RegExp('APP' + '_PIN\\s*[:=]\\s*[\'"][^\'"]+[\'"]')],
  ];
  const hits = [];
  for (const [f, t] of texts) for (const [label, re] of patterns) {
    const m = t.match(re);
    if (m) hits.push(`${f}:${lineOf(t, m.index)} — looks like a ${label}`);
  }
  report(hits.length ? 'FAIL' : 'PASS', 'No secrets in files to be committed', hits);
}

// 8. Leftover merge-conflict markers.
{
  const hits = [];
  for (const [f, t] of texts) for (const m of t.matchAll(/^(<{7}|={7}|>{7})( |$)/gm)) hits.push(`${f}:${lineOf(t, m.index)}`);
  report(hits.length ? 'FAIL' : 'PASS', 'No merge-conflict markers', hits);
}

// 9. External scripts/styles only from the allowed hosts (warning only).
{
  const allowed = /^https:\/\/(cdnjs\.cloudflare\.com|cdn\.jsdelivr\.net|fonts\.googleapis\.com|fonts\.gstatic\.com)(\/|$)/;
  const bad = [...markup.matchAll(/<(?:script|link)\b[^>]*\b(?:src|href)="(https?:[^"]+)"/g)]
    .filter(m => !allowed.test(m[1]))
    .map(m => `index.html:${lineOf(markup, m.index)} — ${m[1]}`);
  report(bad.length ? 'WARN' : 'PASS', 'External resources from allowed hosts', bad);
}

rmSync(tmp, { recursive: true, force: true });
console.log(failed ? `\n${failed} check(s) failed.` : '\nAll checks passed.');
process.exit(failed ? 1 : 0);
