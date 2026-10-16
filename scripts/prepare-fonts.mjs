import { readFile, writeFile } from "node:fs/promises";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const styles = [
  ["Regular", "Regular"],
  ["Bold", "Bold"],
  ["Italic", "Italic"],
  ["BoldItalic", "Bold Italic"],
];

export const FontFaces = [
  ...[
    ["Carlito", "Carlito", "Office Viewer Sans", "Carlito-OFL.txt",
      'Copyright 2013 The Carlito Project Authors (https://github.com/googlefonts/carlito), with Reserved Font Name "Carlito"'],
    ["LiberationSans", "Liberation Sans", "Office Viewer Sans Alt", "Liberation-LICENSE.txt",
      "Digitized data copyright (c) 2010 Google Corporation with Reserved Font Arimo, Tinos and Cousine. Copyright (c) 2012 Red Hat, Inc. with Reserved Font Name Liberation."],
    ["LiberationSerif", "Liberation Serif", "Office Viewer Serif", "Liberation-LICENSE.txt",
      "Digitized data copyright (c) 2010 Google Corporation with Reserved Font Arimo, Tinos and Cousine. Copyright (c) 2012 Red Hat, Inc. with Reserved Font Name Liberation."],
    ["LiberationMono", "Liberation Mono", "Office Viewer Mono", "Liberation-LICENSE.txt",
      "Digitized data copyright (c) 2010 Google Corporation with Reserved Font Arimo, Tinos and Cousine. Copyright (c) 2012 Red Hat, Inc. with Reserved Font Name Liberation."],
  ].flatMap(([stem, family, name, license, copyright]) => styles.map(([suffix, subfamily]) => ({
    file: `${stem}-${suffix}.ttf`, family, name, subfamily, license, copyright,
    postscript: `${name.replaceAll(" ", "")}-${suffix}`,
  }))),
  {
    file: "NotoSansSC-Regular.ttf", family: "Noto Sans SC", name: "Office Viewer CJK", subfamily: "Regular",
    postscript: "OfficeViewerCJK-Regular", license: "NotoSansSC-OFL.txt",
    copyright: "Copyright 2014-2021 Adobe (http://www.adobe.com/), with Reserved Font Name 'Source'",
  },
];

const requiredTables = ["cmap", "head", "hhea", "hmtx", "maxp", "OS/2", "post", "glyf", "loca"];
const fontChecksum = 0xB1B0AFBA;
const align4 = size => Math.ceil(size / 4) * 4;

export function checksum(bytes) {
  let sum = 0;
  let offset = 0;
  for (; offset + 4 <= bytes.length; offset += 4) {
    sum = (sum + bytes.readUInt32BE(offset)) >>> 0;
  }
  let tail = 0;
  for (let index = 0; offset + index < bytes.length; index++) {
    tail |= bytes[offset + index] << (24 - index * 8);
  }
  return (sum + (tail >>> 0)) >>> 0;
}

export function fontTables(bytes) {
  if (bytes.length < 12 || bytes.readUInt32BE(0) !== 0x00010000) {
    throw new Error("Expected a single TrueType font");
  }
  const count = bytes.readUInt16BE(4);
  const directoryEnd = 12 + count * 16;
  if (!count || directoryEnd > bytes.length) throw new Error("Invalid font directory");
  const tables = new Map();
  for (let index = 0; index < count; index++) {
    const record = 12 + index * 16;
    const tag = bytes.toString("ascii", record, record + 4);
    const offset = bytes.readUInt32BE(record + 8);
    const length = bytes.readUInt32BE(record + 12);
    if (tables.has(tag) || offset % 4 || offset < directoryEnd || offset + length > bytes.length) {
      throw new Error(`Invalid font table: ${tag}`);
    }
    tables.set(tag, bytes.subarray(offset, offset + length));
  }
  for (const tag of requiredTables) {
    if (!tables.has(tag)) throw new Error(`Missing font table: ${tag}`);
  }
  if (tables.get("head").length < 54) throw new Error("Invalid font header");
  return tables;
}

function namingTable(face, revision) {
  const version = revision.toFixed(3);
  const names = [
    [0, face.copyright],
    [1, face.name],
    [2, face.subfamily],
    [3, `${version};OfficeViewer;${face.postscript}`],
    [4, `${face.name} ${face.subfamily}`],
    [5, `Version ${version}`],
    [6, face.postscript],
    [10, `Modified from ${face.family} font programs for Office Viewer. Glyphs and metrics are unchanged; naming metadata and checksums are rebuilt.`],
    [13, `Licensed under the SIL Open Font License 1.1. See LICENSES/${face.license} for the full license, original copyright and Reserved Font Name notices.`],
    [14, "https://openfontlicense.org/open-font-license-official-text/"],
  ];
  const strings = names.map(([, text]) => Buffer.from(text, "utf16le").swap16());
  const storage = 6 + names.length * 12;
  const bytes = Buffer.alloc(storage + strings.reduce((size, text) => size + text.length, 0));
  bytes.writeUInt16BE(names.length, 2);
  bytes.writeUInt16BE(storage, 4);
  let offset = 0;
  names.forEach(([id], index) => {
    const record = 6 + index * 12;
    bytes.writeUInt16BE(3, record); // Windows
    bytes.writeUInt16BE(1, record + 2); // Unicode BMP, UTF-16BE
    bytes.writeUInt16BE(0x0409, record + 4); // English (US)
    bytes.writeUInt16BE(id, record + 6);
    bytes.writeUInt16BE(strings[index].length, record + 8);
    bytes.writeUInt16BE(offset, record + 10);
    strings[index].copy(bytes, storage + offset);
    offset += strings[index].length;
  });
  return bytes;
}

export function validateFont(bytes, face) {
  const tables = fontTables(bytes);
  const names = tables.get("name");
  if (!names || names.length < 6) throw new Error(`${face.file}: missing name table`);
  const count = names.readUInt16BE(2);
  const storage = names.readUInt16BE(4);
  if (names.readUInt16BE(0) !== 0 || storage < 6 + count * 12 || storage > names.length) {
    throw new Error(`${face.file}: invalid name table`);
  }
  const values = new Map();
  for (let index = 0; index < count; index++) {
    const record = 6 + index * 12;
    const length = names.readUInt16BE(record + 8);
    const offset = storage + names.readUInt16BE(record + 10);
    if (length % 2 || offset + length > names.length) throw new Error(`${face.file}: invalid font name`);
    if (names.readUInt16BE(record) === 3 && names.readUInt16BE(record + 2) === 1 && names.readUInt16BE(record + 4) === 0x0409) {
      values.set(names.readUInt16BE(record + 6), Buffer.from(names.subarray(offset, offset + length)).swap16().toString("utf16le"));
    }
  }
  if (values.get(0) !== face.copyright || values.get(1) !== face.name || values.get(2) !== face.subfamily || values.get(4) !== `${face.name} ${face.subfamily}` || values.get(6) !== face.postscript || !values.get(13)?.includes("SIL Open Font License 1.1")) {
    throw new Error(`${face.file}: font naming does not match the exported face`);
  }
  if (checksum(bytes) !== fontChecksum) throw new Error(`${face.file}: invalid font checksum`);
  for (let index = 0; index < bytes.readUInt16BE(4); index++) {
    const record = 12 + index * 16;
    const tag = bytes.toString("ascii", record, record + 4);
    const table = Buffer.from(tables.get(tag));
    if (tag === "head") table.writeUInt32BE(0, 8);
    if (checksum(table) !== bytes.readUInt32BE(record + 4)) throw new Error(`${face.file}: invalid ${tag} checksum`);
  }
}

// PDF font programs omit naming metadata. Use distinct names for the modified
// OFL fonts, retaining the original notices and font programs. CSS aliases and
// filenames remain lookup keys for the existing pagelayout substitution map.
export function prepareFont(bytes, face) {
  const tables = fontTables(bytes);
  const head = Buffer.from(tables.get("head"));
  const names = namingTable(face, head.readInt32BE(4) / 65536);
  if (tables.get("name")?.equals(names)) {
    validateFont(bytes, face);
    return bytes;
  }
  head.writeUInt32BE(0, 8);
  tables.set("head", head);
  tables.set("name", names);
  const entries = [...tables].sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0);
  const count = entries.length;
  let offset = 12 + count * 16;
  const output = Buffer.alloc(offset + entries.reduce((size, [, table]) => size + align4(table.length), 0));
  output.writeUInt32BE(0x00010000, 0);
  output.writeUInt16BE(count, 4);
  const exponent = Math.floor(Math.log2(count));
  const searchRange = 2 ** exponent * 16;
  output.writeUInt16BE(searchRange, 6);
  output.writeUInt16BE(exponent, 8);
  output.writeUInt16BE(count * 16 - searchRange, 10);
  let headOffset;
  entries.forEach(([tag, table], index) => {
    const record = 12 + index * 16;
    output.write(tag, record, 4, "ascii");
    output.writeUInt32BE(checksum(table), record + 4);
    output.writeUInt32BE(offset, record + 8);
    output.writeUInt32BE(table.length, record + 12);
    table.copy(output, offset);
    if (tag === "head") headOffset = offset;
    offset += align4(table.length);
  });
  output.writeUInt32BE((fontChecksum - checksum(output)) >>> 0, headOffset + 8);
  validateFont(output, face);
  return output;
}

async function main() {
  const args = process.argv.slice(2);
  const check = args.includes("--check");
  const directories = args.filter(arg => arg !== "--check");
  if (directories.length > 1 || directories.some(arg => arg.startsWith("--"))) {
    throw new Error("Usage: node scripts/prepare-fonts.mjs [font-directory] [--check]");
  }
  const directory = directories[0] ? resolve(directories[0]) : resolve(dirname(fileURLToPath(import.meta.url)), "../public/fonts");
  let changed = 0;
  for (const face of FontFaces) {
    if (!(await readFile(join(directory, "LICENSES", face.license))).length) throw new Error(`Missing license: ${face.license}`);
    const path = join(directory, face.file);
    const bytes = await readFile(path);
    if (check) validateFont(bytes, face);
    else {
      const prepared = prepareFont(bytes, face);
      if (!prepared.equals(bytes)) { await writeFile(path, prepared); changed++; }
    }
  }
  console.log(`${check ? "Checked" : "Prepared"} ${FontFaces.length} browser fonts${check ? "" : ` (${changed} updated)`}`);
}

if (process.argv[1] && import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  main().catch(error => { console.error(error.message); process.exitCode = 1; });
}
