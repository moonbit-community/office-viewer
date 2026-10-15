import { readFile } from 'node:fs/promises';

// Built by the npm test scripts, so assertions exercise the production MoonBit
// package rather than extracting JavaScript strings from its source files.
export const runtime = await readFile(new URL(
  '../_build/js/debug/build/tools/runtime-harness/runtime-harness.js', import.meta.url,
), 'utf8');
