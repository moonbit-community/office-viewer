import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';

// Built by the npm test scripts, so assertions exercise the production MoonBit
// package rather than extracting JavaScript strings from its source files.
const candidates = [
  '../_build/js/debug/build/tools/runtime-harness/runtime-harness.js',
  '../_build/js/debug/build/Lampese/office-viewer/tools/runtime-harness/runtime-harness.js',
  '../_build/js/debug/build/moonbit-community/office-viewer/tools/runtime-harness/runtime-harness.js',
].map(path => new URL(path, import.meta.url));
export const runtime = await readFile(candidates.find(path => existsSync(path)) || candidates[0], 'utf8');
