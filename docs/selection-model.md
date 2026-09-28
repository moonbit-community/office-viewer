# Office Selection Model

This document defines the boundary between the read-only Office renderer and
the later OpenSeek Agent integration.

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

## Layers

The implementation is deliberately split into three layers:

1. The document layer reads OOXML and renders through the unmodified
   `office.mbt` layout packages.
2. The browser layer adds interaction metadata that can be proved from the
   rendered DOM: Excel cell coordinates and Word physical selection context.
3. The host receives the versioned envelope and can use canonical selectors
   only where the renderer actually has one.

Word fragment spans are local to the rendered DOM text node and use UTF-16
offsets, matching the DOM and JavaScript string model. They are not byte
offsets into XML and are not a substitute for a canonical selector. Byte or
OOXML path metadata can be added later by a separate source-map-capable
renderer without changing the selection envelope.

## Current milestone

The current implementation covers native Word text selection and Excel range
selection over the rendered worksheet window. Nested tables, drawings, and
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

Each returns `office.selection.materialized.v1` and keeps the original
`office.selection.v1` shape under `selection` for the fields that are safe for
the selected mode. The `content` object makes omission explicit. `full` keeps
the exact selected text and original selection details. `preview` uses a
deterministic UTF-16-bounded head/tail excerpt and does not invent a summary;
the `[...]` marker is not counted as selected document text. `deferred` sets
`selection.text` to `null` and carries compact reader references instead; the
host can use the `file` and `selectors` to read the exact text later.

Preview and deferred results intentionally omit large DOM rectangles, raw Word
anchor text, and Excel cell payloads from the projected selection. They retain
counts and compact source fields (`selector`, `path`, `part`, UTF-16 spans,
`runs`, `page`, `stability`, and reader judgment fields). The convenience APIs
return at most 64 compact source anchors and report `anchors_truncated`; callers
that need another bound can use `materialize_selection` with
`SelectionMaterializationOptions`.

For Excel selections that do not have a top-level `text` field, the viewer
derives exact content from the selected cells in row order, using tabs between
columns and newlines between rows. The canonical sheet/range selector remains
the authoritative read-back address.

## OpenSeek handoff

OpenSeek should receive the envelope together with the `file` path and a
host-owned document identity when one is available. The standalone demo uses a
`/mock/office/...` path; the OpenSeek host replaces it with the path it passes
to `office`. `selector` follows `office.selector/1`; `path`, `part`, `runs`,
`stability`, and the reader judgment fields are bounded source metadata for
debugging and later Agent actions. OpenSeek can quote `text` immediately and
use a canonical selector to request richer context through `office get`,
`office text`, `office outline`, or `office query` later. This keeps Agent
behavior independent from the current SVG/HTML implementation.
