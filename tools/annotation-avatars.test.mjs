import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import vm from "node:vm";

// Execute the production FFI helper. Browser verification covers its mounting
// in card author rows; these cases cover real document-author metadata.
const source = await readFile(new URL("../viewer/ffi.mbt", import.meta.url), "utf8");
const runtime = source.split("\n").filter(line => line.startsWith("  #|"))
  .map(line => line.slice(5)).join("\n");
const start = runtime.indexOf("const avatarSegmenter =");
const end = runtime.indexOf("// Avatar presentation travels", start);
assert.ok(start >= 0 && end > start, "annotation avatar runtime is present");

function element(tag, className = "", text = "") {
  const node = {
    tag, className, textContent: text, children: [], attributes: {}, properties: {},
    setAttribute(name, value) { this.attributes[name] = value; },
    appendChild(child) { this.children.push(child); },
  };
  node.style = { setProperty(name, value) { node.properties[name] = value; } };
  return node;
}

function factory(internationalization = Intl) {
  const document = {
    createElementNS(namespaceURI, tag) { return { ...element(tag), namespaceURI }; },
  };
  return vm.runInNewContext(
    `${runtime.slice(start, end)}; annotationAvatar`,
    { document, element, Intl: internationalization },
  );
}
const avatar = factory();

test("document initials take priority and long values stay compact", () => {
  assert.equal(avatar({ author: "Ada Lovelace", initials: " xy " }).textContent, "XY");
  assert.equal(avatar({ author: "Ada Lovelace", initials: "User" }).textContent, "US");
  assert.equal(avatar({ author: null, initials: "AL" }).textContent, "AL");
});

test("names support Latin words, Chinese names, and complete Unicode graphemes", () => {
  for (const [author, expected] of [
    ["Ada Lovelace", "AL"], ["  Ada   Byron Lovelace  ", "AL"],
    ["产品评审", "产品"], ["𠮷野", "𠮷野"], ["e\u0301mile", "ÉM"],
  ]) assert.equal(avatar({ author, initials: null }).textContent, expected);
  assert.equal(avatar({ author: "Engineer", initials: "👩🏽‍💻A" }).textContent, "👩🏽‍💻A");
});

test("an author's color stays stable across comments, replies, and initials", () => {
  const original = avatar({ author: "Ada Lovelace", initials: "AL" });
  assert.deepEqual(avatar({ author: "Ada Lovelace", initials: null }).properties, original.properties);
  assert.deepEqual(avatar({ author: " Ada Lovelace ", initials: "A" }).properties, original.properties);
  const backgrounds = ["Ada Lovelace", "Bob", "陈明", "Lee", "产品评审"]
    .map(author => avatar({ author }).properties["--office-avatar-background"]);
  assert.ok(new Set(backgrounds).size > 1);
  assert.deepEqual(
    avatar({ author: "e\u0301mile", initials: null }).properties,
    avatar({ author: "émile", initials: null }).properties,
  );
});

test("missing or blank metadata produces a decorative person icon", () => {
  for (const item of [{}, { author: null, initials: null }, { author: " \n ", initials: "\t" }]) {
    const result = avatar(item);
    assert.equal(result.textContent, "");
    assert.equal(result.attributes["aria-hidden"], "true");
    assert.equal(result.children.length, 1);
    assert.equal(result.children[0].tag, "svg");
    assert.deepEqual(result.children[0].children.map(node => node.tag), ["circle", "path"]);
  }
});

test("author-controlled text stays text and creates no remote image", () => {
  const result = avatar({ author: "<img src=https://example.com/avatar>", initials: "<b>" });
  assert.equal(result.textContent, "<B");
  assert.deepEqual(result.children, []);
  assert.equal(result.attributes["aria-hidden"], "true");
});

test("browsers without grapheme segmentation preserve supplementary characters", () => {
  const fallback = factory({ Segmenter: undefined });
  assert.equal(fallback({ author: "𠮷野", initials: null }).textContent, "𠮷野");
  assert.equal(fallback({ author: null, initials: null }).children[0].tag, "svg");
});
