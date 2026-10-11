# Office Selection Model

This document defines the selections emitted by the read-only document surface
and consumed by host applications and agents.

## Goals

The viewer must preserve the native interaction of each document type while
returning a bounded selection context. Agent code should not parse SVG, HTML,
or browser-specific selection objects.

The selection envelope is versioned and contains both the text the user saw
and anchors back to the source document. A renderer may add fields, but it must
not change the meaning of existing fields.

```json
{
  "schema": "office.selection.v1",
  "file": "/mock/office/report.docx",
  "format": "docx",
  "kind": "word",
  "text": "selected text",
  "anchors": [
    {
      "stability": "physical-only",
      "start_utf16": 4,
      "end_utf16": 17,
      "page": 2,
      "text": "selected text",
      "rect": { "left": 120, "top": 240, "width": 84, "height": 16 }
    }
  ],
  "pages": [2]
}
```

Word continues to use browser-native text selection. The browser intersects the
native `Range` with positioned SVG text fragments and records the selected text,
page number, UTF-16 span within each visual fragment, and a physical bounding
rectangle. Before emitting the SVG, the viewer reads the same DOCX bytes with
`office.mbt`'s annotated reader, builds the body paragraph/run projection, and
joins it to the PageModel text in document order. A fragment receives
`data-source-*` attributes only after the paragraph projection, parsed run
text, and reader `find` run paths agree. The browser then translates the
fragment-local range to paragraph-relative UTF-16 offsets.

The join is deliberately fail-closed. Unsupported wrappers, numbering or
headers/footers that make the body projection differ from the layout, duplicate
or missing anchors, and any incomplete text sequence retain the physical page
rectangle and omit a logical selector. A physical-only anchor is intentionally
less powerful than a guessed selector and must be treated as such by the Agent.

For a joined body paragraph, metadata follows the existing `office.mbt` reader
contract: `path` is the reader path such as `p[3]`, `selector` is
`/docx/body/p[3]` or a unique `p[id="..."]` selector, `part` is
`word/document.xml`, `start_utf16`/`end_utf16` are paragraph projection
coordinates, and `runs` contains body-relative run paths such as
`p[3]/r[2]`. The reader's `actionable`, refusal reason, and paragraph anchor
status are carried as optional metadata so a later Agent can distinguish a
valid address from a reader refusal without reparsing SVG.

Excel uses cell selection rather than text selection:

```json
{
  "schema": "office.selection.v1",
  "file": "/mock/office/budget.xlsx",
  "format": "xlsx",
  "kind": "excel",
  "sheet": "Sheet1",
  "range": "B4:D9",
  "selector": "/xlsx/sheet[name=\"Sheet1\"]/range[B4:D9]",
  "stability": "snapshot-relative",
  "cells": [
    { "address": "B4", "row": 4, "column": 2, "text": "..." }
  ]
}
```

The upstream `xlsx2html` output already carries worksheet sections and table
structure but does not carry cell coordinates. The browser adds row, column,
sheet, and A1 address attributes from the actual rendered table, accounting for
rowspan and colspan. Merged cells retain the anchor cell's address. The emitted
Excel selector therefore follows the canonical `/xlsx/sheet[name=...]/cell[...]`
shape for one cell and `/xlsx/sheet[name=...]/range[...]` for a rectangle,
matching `office.mbt`'s selector parser.

## Annotation-body selections

Comments and Notes are separate selectable content, not document text or a cell
value. Clicking **Select comment** selects the complete available body; dragging
within one contextual comment body selects just that substring. Both paths use
`kind: "annotation"` in the existing envelope:

```json
{
  "schema": "office.selection.v1",
  "kind": "annotation",
  "format": "docx",
  "file": "/docs/report.docx",
  "document": {"file": "/docs/report.docx", "sha256": "<source-byte SHA-256>"},
  "text": "selected words",
  "annotation": {
    "id": "docx.comment:0",
    "type": "docx.comment",
    "source_id": "17",
    "author": "Ada",
    "parent_id": null,
    "resolved": false,
    "anchor": {
      "kind": "docx",
      "granularity": "paragraph-context",
      "confidence": "source-node",
      "locations": [{"part": "word/document.xml", "path": "p[3]"}],
      "ranges": []
    }
  },
  "annotation_range": {
    "scope": "text",
    "start_utf16": 5,
    "end_utf16": 19,
    "body_utf16_length": 40
  },
  "anchors": []
}
```

The half-open UTF-16 span is relative to the plain annotation body returned in
`office.annotations.v1`, including its explicit line breaks and paragraph
separators. It is never a span in the Word paragraph or Excel cell. A `whole`
selection covers `[0, body_utf16_length)`. Cross-card and body-to-document ranges
are not emitted as mixed selections. Missing or ambiguous bodies cannot be
selected. Unknown author/thread/resolution metadata stays missing or null;
`false` still means an explicitly open comment.

The annotation reference retains the original type, source ID, author/thread
fields, source anchors, and anchor confidence, but omits full body text and
`body_candidates`. Only top-level `text` contains selected content. A Note uses
`format: "xlsx"`, `annotation.type: "xlsx.note"`, and its original sheet/cell
anchor. Off-preview Notes and unmapped Word comments can still be selected from
the list without claiming a rendered location.

Hosts can create the same envelope without DOM selection:

```mbt
@viewer.annotation_selection(item) // whole body
@viewer.annotation_selection(item, start_utf16=5, end_utf16=19, scope="text")
```

The constructor rejects out-of-bounds spans and surrogate-pair splits. Item IDs
are snapshot-local: always retain and match the source file plus SHA-256, not
just an ID. This constructor projects metadata; it does not reread or authenticate
a host-supplied item. Browser commands resolve the requested item against the
currently installed snapshot before selecting it.

All three materialization modes support this kind. Preview bounds only selected
text; deferred removes it entirely while retaining the annotation identity and
body range. Neither mode embeds the unselected body. Their source reference has
`reference_kind: "annotation"`, `annotation_schema: "office.annotations.v1"`,
`resolve_with: "read_annotations"`, `selector_schema: null`, and `readable: false`.
Here `readable` means readable through the existing document selector route,
which does not address comment bodies. To resolve, re-extract annotations from
the same source-byte hash, find the item ID, and apply the recorded body span.
Never substitute the surrounding paragraph/cell text. Consumers must handle the
new kind explicitly instead of assuming every selection is document content.

## PDF selections

PDF uses native text selection over PDF.js's text layer. The envelope has
`format: "pdf"`, `kind: "pdf"`, `file`, `text`, `browser_text`, `document`,
`reference`, `anchors`, `pages`, and `page_info`. Its authoritative reference
is independent of `office.selector/1`:

```json
{
  "schema": "office.pdf.text-range.v1",
  "file": "/mock/office/sample.pdf",
  "document": {
    "sha256": "26d937576b2e771cbd616d8dac91a92c5d500b17cdedabdf17412dbc01b121f2",
    "byte_length": 9639,
    "page_count": 4,
    "parser": { "name": "pdfjs", "version": "6.4.299" }
  },
  "text_content": {
    "include_marked_content": true,
    "disable_normalization": true
  },
  "join": "pdfjs-items-eol-v1",
  "start": { "page": 1, "item_index": 1, "offset_utf16": 0 },
  "end": { "page": 1, "item_index": 1, "offset_utf16": 13 },
  "text_utf16_length": 13
}
```

Pages are one-based. Item indices are zero-based indices into the original
`getTextContent().items` array, including marked-content entries. Offsets are
UTF-16 positions within a text item's original string; the end is exclusive.
Indices are relative to identical bytes, the parser version, and the recorded
extraction options. They are not byte offsets or persistent PDF object ids.
`document` is also present at the top level of the raw selection and may
include PDF.js fingerprints; SHA-256 is the content identity.

`pdfjs-items-eol-v1` preserves item strings, inserts a newline after `hasEOL`
items except the final item, and joins pages with two newlines. `browser_text`
retains the browser's native selected string separately. The runtime verifies
the rendered and extracted item sequences before mapping them, and does not
invent a logical address when they disagree.

Each anchor records its page, page object reference, item index, selected UTF-16
span, font, direction, and geometry. Geometry has `space:
"pdf-page-user-space"`, `unit: "user-unit"`, the page's `user_unit`, quads, and
a bounding box. `selected-dom-range` precision describes a captured DOM range
converted through the page viewport. `whole-text-item` precision describes
metric bounds for the entire source item during resolution without a DOM range;
it does not provide per-character geometry. `page_info` retains the CropBox,
intrinsic page rotation, user unit, and presence of a structure tree.

Available tagged structure adds `marked_content_id` and `semantics` containing
the actual `role_path`, inherited language, and alternate text. Untagged PDFs
have no inferred paragraph or table roles. Image-only pages have no text
anchors, and OCR is not included.

## Layers

The implementation is deliberately split into three layers:

1. The document layer renders OOXML through unmodified `office.mbt` packages
   and PDF through PDF.js.
2. The browser layer adds interaction metadata that can be proved from the
   rendered DOM and source readers: Excel cell coordinates, validated Word
   paragraph/run joins, and PDF text-item ranges.
3. The host receives the versioned envelope and can use canonical selectors
   only where the renderer actually has one.

Word spans use paragraph-relative UTF-16 offsets when the source join succeeds
and fragment-local offsets for physical-only anchors. Neither is a byte offset
into XML. Hosts must inspect stability and selector availability before using
a source address.

## Current milestone

The current implementation covers native Word and PDF text selection and Excel
range selection over the rendered worksheet window. Nested tables, drawings, and
other unsupported OOXML remain governed by the layout engine's existing
diagnostics instead of being treated as selected content.

## Agent materialization

The viewer package does not force one context policy on its host. It exposes
three JSON APIs over the existing envelope:

```mbt
@viewer.materialize_selection_full(selection_json)
@viewer.materialize_selection_preview(selection_json, 1000)
@viewer.materialize_selection_deferred(selection_json)
```

Each returns `office.selection.materialized.v1` and keeps selected content under
`selection`. The `content` object makes omission explicit. `full` keeps the exact
selected text; `preview` uses a deterministic UTF-16-bounded head/tail excerpt
and does not invent a summary; the `[...]` marker is not counted as selected
document text. `deferred` sets `selection.text` to `null` and carries a source
reference so the host can read the exact text later.

PDF materializations deliberately use a compact projection. `selection` contains
only `schema`, `format`, `kind`, and the mode's `text`. `source` contains
`readable` and one complete `reference` with the file path, document identity,
parser/extraction settings, and exact text endpoints. It does not repeat that
reference or identity, and omits raw browser text, per-item anchors, geometry,
tagged semantics, page metadata, and generic selector/structure counts. An
invalid reference is omitted and `source.reason` explains why it cannot be
resolved. The PDF item range is not truncated by the compact projection.
`resolve_pdf_selection` accepts the materialized value and verifies the active
bytes and PDF.js version before reading its text items again.

The original PDF selection callback and resolved PDF selection retain the
detailed `anchors`, `pages`, and `page_info` data for applications that need
geometry or tagged structure. Compact Agent materialization does not remove that
data from the viewer's internal selection or resolver result.

For DOCX/XLSX, preview and deferred results omit large DOM rectangles, raw Word
anchor text, and Excel cell payloads from the projected selection. They retain
counts and compact source fields (`selector`, `path`, `part`, UTF-16 spans,
`runs`, `page`, `stability`, and reader judgment fields). The convenience APIs
return at most 64 compact source anchors and report `anchors_truncated`; callers
that need another bound can use `materialize_selection` with
`SelectionMaterializationOptions`.

Detailed PDF anchors retain item indices, page references, direction, font,
marked-content ids, tagged semantics, and geometry. `resolve_pdf_selection`
accepts any of the three materialized forms and reconstructs full text by
rereading PDF.js items. It returns whole-item geometry when no live DOM range
is supplied.

For Excel selections that do not have a top-level `text` field, the viewer
derives exact content from the selected cells in row order, using tabs between
columns and newlines between rows. The canonical sheet/range selector remains
the authoritative read-back address.

## OpenSeek handoff

OpenSeek should receive the envelope together with the `file` path and a
host-owned document identity when one is available. The standalone demo uses a
`/mock/office/...` path; the OpenSeek host replaces it with the path it passes
to `office`. For DOCX/XLSX, `selector` follows `office.selector/1`; `path`, `part`, `runs`,
`stability`, and the reader judgment fields are bounded source metadata for
debugging and later Agent actions. OpenSeek can quote `text` immediately and
use a canonical selector to request richer context through `office get`,
`office text`, `office outline`, or `office query` later. This keeps Agent
behavior independent from the current SVG/HTML implementation.

PDF uses the same host-owned file-path field but an explicit PDF.js reference
and resolver. It does not claim that `office.mbt` or `pdflite` can consume the
PDF.js item indices as selectors. A host may add another parser adapter later
without changing the meaning of these existing references.
