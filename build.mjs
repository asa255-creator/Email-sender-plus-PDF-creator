#!/usr/bin/env node
/**
 * build.mjs
 * Concatenates every src/*.gs module into the single dist/Code.gs file that
 * gets pasted into the Apps Script editor.
 *
 * Apps Script already shares one global scope across all .gs files in a
 * project, so combining them changes nothing about how the code runs.
 *
 * Usage: node build.mjs [--check]
 *   --check  exit non-zero if dist/Code.gs is stale (used by CI)
 */

import { readFileSync, writeFileSync, readdirSync, mkdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';

const SRC_DIR = 'src';
const OUT_FILE = join('dist', 'Code.gs');

// Explicit order keeps the bundle readable and the diff stable. Anything not
// listed is appended alphabetically, so adding a module never silently breaks
// the build.
const ORDER = [
  'OPEN.gs',
  'EmailSender.gs',
  'BounceChecker.gs',
  'EmailFinderGoogle.gs',
  'EmailFinderVCF.gs',
  'PDFBundler.gs',
  'Utilities.gs'
];

function orderedModules() {
  const found = readdirSync(SRC_DIR).filter(f => f.endsWith('.gs'));
  const known = ORDER.filter(f => found.includes(f));
  const extra = found.filter(f => !ORDER.includes(f)).sort();
  const missing = ORDER.filter(f => !found.includes(f));

  if (missing.length) {
    console.warn(`warning: listed in ORDER but not found in ${SRC_DIR}/: ${missing.join(', ')}`);
  }
  return [...known, ...extra];
}

function banner(text) {
  const line = '/'.repeat(76);
  return `${line}\n// ${text}\n${line}`;
}

function build() {
  const modules = orderedModules();

  const header = [
    banner('GENERATED FILE - DO NOT EDIT'),
    '//',
    '// This is every file in src/ concatenated into one, so it can be pasted',
    '// into the Apps Script editor in a single step.',
    '//',
    '// Edit the modules in src/ and run `node build.mjs` instead. CI regenerates',
    '// this file on every push, so hand edits here will be overwritten.',
    '//',
    `// Modules: ${modules.join(', ')}`,
    banner('END HEADER'),
    ''
  ].join('\n');

  const body = modules.map(name => {
    const source = readFileSync(join(SRC_DIR, name), 'utf8').replace(/\s+$/, '');
    return `\n${banner(`src/${name}`)}\n\n${source}\n`;
  }).join('\n');

  return `${header}${body}`;
}

const output = build();

if (process.argv.includes('--check')) {
  const current = existsSync(OUT_FILE) ? readFileSync(OUT_FILE, 'utf8') : '';
  if (current !== output) {
    console.error(`${OUT_FILE} is out of date. Run: node build.mjs`);
    process.exit(1);
  }
  console.log(`${OUT_FILE} is up to date.`);
  process.exit(0);
}

mkdirSync('dist', { recursive: true });
writeFileSync(OUT_FILE, output);

const lines = output.split('\n').length;
console.log(`Wrote ${OUT_FILE} (${lines} lines, ${(output.length / 1024).toFixed(1)} KB)`);
