# Embedding the viewer surface

The `moonbit-community/office-viewer/viewer` package is the reusable part of
the project. It is a JS/Rabbita component package, while the bytes and the
canonical file path remain owned by the host application.

An embedded host follows this sequence:

1. The native host reads a DOCX or XLSX file and keeps the path that it passes
   to `office.mbt`.
2. The host transfers the bytes and path to its JS frontend.
3. The frontend calls `render_document(name, file, bytes)` and stores the
   returned `RenderedDocument` in its own state.
4. The frontend renders `render_surface(SurfaceConfig { ... })` inside its
   existing Rabbita layout.
5. After the state update, it returns `interactions_command(config, callback)`
   and `visual_command(config)` from the update command batch.

`SurfaceConfig.surface_id` is host-configurable, so multiple viewers can be
mounted without sharing DOM ids. `document_file` is passed unchanged into the
selection envelope. The standalone demo uses `/mock/office/<name>`; OpenSeek
must use the real path understood by its native `office.mbt` command layer.

Excel selection emits a canonical `office.selector/1` address. A single cell
uses `/xlsx/sheet[name="Data"]/cell[B4]`; a rectangular drag uses
`/xlsx/sheet[name="Data"]/range[B4:D9]`. The browser derives these addresses
from the rendered table, including merged cell spans.

The host chooses how much selected content to send to an Agent by calling one
of the viewer's materialization APIs after receiving the callback JSON:

```mbt
let full = @viewer.materialize_selection_full(selection_json)
let preview = @viewer.materialize_selection_preview(selection_json, 1000)
let deferred = @viewer.materialize_selection_deferred(selection_json)
```

All three return `office.selection.materialized.v1`. `full` carries exact text,
`preview` carries a deterministic head/tail excerpt and UTF-16 omission counts,
and `deferred` carries `file`, canonical selectors, and compact source anchors
without the selected text. The generic `materialize_selection` entry point
accepts `SelectionMaterializationOptions` when the host needs a different
preview or source-anchor bound. A deferred result is a reference for a later
`office.mbt` read; it is not an instruction to guess text from the rendered
SVG or HTML.

The clean upstream `pagelayout/svg` backend currently exposes no OOXML source
map. Word selection therefore emits the selected text, physical page, UTF-16
fragment spans, and a bounding rectangle with `stability: "physical-only"`.
It omits a canonical selector until a renderer supplies a source map. This
keeps the Agent handoff honest and lets a future provenance-capable
`office.mbt` renderer improve the same envelope without changing the viewer
mounting API.
