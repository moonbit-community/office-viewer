import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import { FontFaces, prepareFont, validateFont } from "./prepare-fonts.mjs";

const font = face => readFileSync(new URL(`../public/fonts/${face.file}`, import.meta.url));

// Read tables independently of the preparer so a directory-writing bug cannot
// make the preservation assertion pass by changing both sides of the check.
function tables(bytes) {
  const result = new Map();
  for (let index = 0; index < bytes.readUInt16BE(4); index++) {
    const record = 12 + index * 16;
    const tag = bytes.toString("ascii", record, record + 4);
    const offset = bytes.readUInt32BE(record + 8);
    result.set(tag, Buffer.from(bytes.subarray(offset, offset + bytes.readUInt32BE(record + 12))));
  }
  return result;
}

function withoutNames(bytes) {
  if (!tables(bytes).has("name")) return Buffer.from(bytes);
  const result = Buffer.from(bytes);
  const count = bytes.readUInt16BE(4);
  let kept = 0;
  for (let index = 0; index < count; index++) {
    const record = 12 + index * 16;
    if (bytes.toString("ascii", record, record + 4) !== "name") {
      bytes.copy(result, 12 + kept++ * 16, record, record + 16);
    }
  }
  result.writeUInt16BE(kept, 4);
  return result;
}

test("all bundled faces contain browser naming metadata, valid checksums and licenses", () => {
  assert.equal(FontFaces.length, 17);
  for (const face of FontFaces) {
    const bytes = font(face);
    assert.ok(tables(bytes).has("name"), `${face.file} must include its name table`);
    validateFont(bytes, face);
    assert.ok(!/Carlito|Liberation|Arimo|Tinos|Cousine|Source|Noto/.test(face.name), "modified font names must not reuse upstream names");
    assert.ok(readFileSync(new URL(`../public/fonts/LICENSES/${face.license}`, import.meta.url)).length);
  }
});

test("preparing PDF font programs preserves every glyph and metric table", () => {
  for (const face of FontFaces) {
    const input = withoutNames(font(face));
    const original = tables(input);
    const prepared = prepareFont(input, face);
    const output = tables(prepared);
    validateFont(prepared, face);
    assert.equal(output.size, original.size + 1);
    for (const [tag, data] of original) {
      const actual = Buffer.from(output.get(tag));
      if (tag === "head") { data.writeUInt32BE(0, 8); actual.writeUInt32BE(0, 8); }
      assert.deepEqual(actual, data, `${face.file}: ${tag} must remain unchanged`);
    }
    assert.deepEqual(prepareFont(prepared, face), prepared, "repeated preparation must not change the font");
  }
});

test("invalid font data and corrupted glyph programs fail validation", () => {
  assert.throws(() => prepareFont(Buffer.alloc(8), FontFaces[0]), /TrueType/);
  const corrupted = prepareFont(withoutNames(font(FontFaces[0])), FontFaces[0]);
  for (let index = 0; index < corrupted.readUInt16BE(4); index++) {
    const record = 12 + index * 16;
    if (corrupted.toString("ascii", record, record + 4) === "glyf") {
      corrupted[corrupted.readUInt32BE(record + 8)] ^= 1;
      break;
    }
  }
  assert.throws(() => validateFont(corrupted, FontFaces[0]), /checksum/);
});
