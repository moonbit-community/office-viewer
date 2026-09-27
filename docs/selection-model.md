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
rectangle. Clean upstream `pagelayout/svg` output does not carry OOXML
provenance attributes, so the viewer omits `selector`, `path`, and source byte
claims unless a future host supplies an explicit source map. A physical-only
anchor is intentionally less powerful than a guessed selector and must be
treated as such by the Agent.

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

## OpenSeek handoff

OpenSeek should receive the envelope together with the `file` path and a
host-owned document identity when one is available. The standalone demo uses a
`/mock/office/...` path; the OpenSeek host replaces it with the path it passes
to `office`. `selector` follows `office.selector/1`, while `source_id` and
`path` remain renderer provenance for debugging. OpenSeek can quote `text`
immediately and use the canonical selector to request richer context through
`office get`, `office text`, `office outline`, or `office query` later. This
keeps Agent behavior independent from the current SVG/HTML implementation.
