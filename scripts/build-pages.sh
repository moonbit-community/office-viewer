#!/usr/bin/env bash
set -euo pipefail

root="$(cd "$(dirname "$0")/.." && pwd)"
dist="$(node -e 'process.stdout.write(require("node:path").resolve(process.argv[1]))' "${1:-$root/dist}")"
# Avoid Warren's unpinned npm-exec fallback.
command -v terser >/dev/null || { echo 'Install terser@5.51.2 and add it to PATH.' >&2; exit 1; }
moonx --target native moonbit-community/warren@0.4.0 \
  -C "$root" build --browser-entry cmd/browser --server-entry "" --dist "$dist"

# Content-addressed assets prevent a cached HTML/JS/CSS version mismatch.
SOURCE_REVISION="${GITHUB_SHA:-$(git -C "$root" rev-parse HEAD)}" node --input-type=module - "$dist" <<'JS'
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, renameSync } from 'node:fs';
import { join } from 'node:path';
const dist = process.argv[2];
let html = readFileSync(join(dist, 'index.html'), 'utf8');
const assets = {};
for (const file of ['index.js', 'styles.css']) {
  const bytes = readFileSync(join(dist, file));
  const sha256 = createHash('sha256').update(bytes).digest('hex');
  const dot = file.lastIndexOf('.');
  const name = `${file.slice(0, dot)}.${sha256.slice(0, 16)}${file.slice(dot)}`;
  const reference = `./${file}`;
  if (html.split(`"${reference}"`).length !== 2) {
    throw new Error(`Expected exactly one HTML reference to ${file}`);
  }
  html = html.replace(`"${reference}"`, `"./${name}"`);
  renameSync(join(dist, file), join(dist, name));
  assets[file] = { file: name, sha256, bytes: bytes.length };
}
writeFileSync(join(dist, 'index.html'), html);
writeFileSync(join(dist, 'build-info.json'), JSON.stringify({
  revision: process.env.SOURCE_REVISION, assets,
}, null, 2) + '\n');
console.log(JSON.stringify({ revision: process.env.SOURCE_REVISION, assets }, null, 2));
JS
