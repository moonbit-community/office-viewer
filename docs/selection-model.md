# Office Selection Model

This document defines the boundary between the read-only Office renderer and
the later OpenSeek Agent integration.

## Goals

The viewer must preserve the native interaction of each document type while
returning a stable, source-aware selection. Agent code should not parse SVG,
HTML, or browser-specific selection objects.

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
      "source_id": "word/document.xml#/w:body/w:p[3]/w:r[2]",
      "part": "word/document.xml",
      "path": "/w:body/w:p[3]/w:r[2]",
      "selector": "/docx/body/p[3]/r[2]",
      "stability": "snapshot-relative",
      "start_utf16": 4,
      "end_utf16": 17,
      "page": 2
    }
  ],
  "pages": [2]
}
```

Word continues to use browser-native text selection. The SVG emitter carries
the source span on each text fragment, and the browser intersects the native
`Range` with those fragments. A selection may contain multiple anchors when it
crosses runs, paragraphs, tables, or pages.

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

Every rendered Excel cell carries its worksheet name, address, row, and column
in `data-*` attributes. Merged cells retain the anchor cell's address.

## Layers

The implementation is deliberately split into three layers:

1. The document layer reads OOXML and assigns stable source paths.
2. The layout layer carries source spans through `RunInput`, `PlacedRun`, and
   `GlyphRun` without changing pagination or coordinates.
3. The browser layer converts native selection into the versioned envelope.

The source span is local to the authored run and uses UTF-16 offsets, matching
the DOM and JavaScript string model. It is not a byte offset into XML. Byte
offsets can be added later for edit operations without changing this selection
contract.

## Current milestone

The first implementation covers body paragraphs, tables, and header/footer
text runs, with SVG metadata, native Word selection, and Excel cell metadata.
Nested tables, drawings, and other unsupported OOXML remain reported by the
layout engine instead of being treated as selected content.

## OpenSeek handoff

OpenSeek should receive the envelope together with the `file` path and a
host-owned document identity when one is available. The standalone demo uses a
`/mock/office/...` path; the OpenSeek host replaces it with the path it passes
to `office`. `selector` follows `office.selector/1`, while `source_id` and
`path` remain renderer provenance for debugging. OpenSeek can quote `text`
immediately and use the canonical selector to request richer context through
`office get`, `office text`, `office outline`, or `office query` later. This
keeps Agent behavior independent from the current SVG/HTML implementation.
