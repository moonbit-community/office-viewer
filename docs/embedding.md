# Embedding the viewer surface

The `moonbit-community/office-viewer/viewer` package provides a JS/Rabbita
surface. The host owns file bytes, the canonical file path, application state,
toolbars, and the policy for sending selections to an agent.

## DOCX and XLSX

1. Read the file bytes and retain the path used by the host's `office.mbt`
   command layer.
2. Call `render_document(name, file, bytes)` and store its `RenderedDocument`.
3. Render `render_surface(config)` inside the existing Rabbita layout.
4. Return `interactions_command(config, on_selection)` and
   `visual_command(config)` after the render.

`SurfaceConfig.surface_id` is host-configurable. Each concurrent surface must
have a unique DOM id. `document_file` is passed unchanged into selections;
the demo uses `/mock/office/<name>` because browser uploads have no native path.

Word source metadata is built within the viewer by reading the same bytes
through `office.mbt`'s annotated reader and joining paragraph/run projections
to the PageModel text. Validated matches carry canonical paragraph selectors
and paragraph-relative UTF-16 spans. Uncertain joins retain physical anchors
without guessed selectors. Excel emits canonical `office.selector/1` cell
and range selectors derived from worksheet coordinates, including merged cells.

## Rendering limits and frozen panes

DOCX surfaces carry the layout reader's limitations banner and
`office.render.diagnostics.v1` JSON. Preserve these elements when embedding.
The JSON reports `complete: false`, counted `items` (`feature`, `count`), and a
`truncated` flag when more than 64 distinct warnings exist. Absence of the
script means no limitations were reported, not proof of complete compatibility.

Excel surfaces contain inert `office-freeze-metadata` JSON with a `sheets` map
of worksheet names to `{ rows, columns }`. `interactions_command` creates each
worksheet's scroll area and installs sticky frozen cells; `visual_command`
refreshes their geometry after zoom. Hosts must run both commands as described
above. `render_document` defaults to `max_rows=1000`, `max_cols=256`; these
positive thresholds choose between static and virtual rendering. They never
truncate the workbook. A total of 32768 cells also triggers virtualization.
The virtual surface retains sparse source data, mounts only the active
worksheet's visible cells, and searches and materializes ranges from the full
data. Notes anywhere in that worksheet remain navigable. The runtime exposes
`surface.__officeRevealExcelCell(sheet, row, column)` for hosts that need to
mount an unloaded citation target before querying its DOM node; a successful
call has already scrolled it into view, so hosts should not apply a subsequent
native `scrollIntoView` to the compressed scroll canvas.
See [rendering support](rendering-support.md).

## PDF

PDF loading is asynchronous. `document_kind(name)` returns `Some(Pdf)` for a
PDF filename; the synchronous `render_document` returns an error directing the
host to `load_pdf_command`.

Render a `SurfaceConfig` with `kind: Some(Pdf)` and
`surface_html: pdf_surface_html()`, then return:

```mbt
@viewer.load_pdf_command(config, bytes, on_selection, on_state)
```

The callbacks return Rabbita commands. Selection events contain
`office.selection.v1` JSON. State events are JSON objects:

| `status` | Fields |
| --- | --- |
| `loading` | `loaded`, `total` |
| `ready` | `pages`, `document` identity |
| `page` | `page` (one-based) |
| `scale` | `zoom` (integer percent) |
| `rotation` | `rotation` (viewer rotation in degrees) |
| `search` | `current`, `total`, optional PDF.js find `state` |
| `password` | `incorrect` |
| `error`, `selection-error` | `message` |

Retain a file generation or request token in asynchronous callbacks and discard
events from an older file. The standalone demo demonstrates this pattern.

Use `visual_command(config)` to update zoom and search. Additional commands are
`pdf_page_command`, `pdf_rotation_command`, `pdf_fit_width_command`,
`pdf_find_command`, and `pdf_password_command`. Rotation is a multiple of 90
degrees. Password values are submitted to the active loading task and are not
included in selection metadata. Cancel a password prompt by disposing the
session.

`pdf_document_info(surface_id)` is asynchronous and returns
`Result[String, String]`. Its `office.pdf.document.v1` JSON contains the file
path, identity, information dictionary, XMP metadata, outline, and permissions
reported by PDF.js. Call it after the document is ready. The demo currently
exposes page navigation rather than an outline sidebar.

## Word and PDF page thumbnails

Render `render_page_thumbnails("page-list", page_label="Page {page}",
unavailable_label="Preview unavailable")` beside the document surface. Keep
its node mounted: the viewer owns only the list's children. The host owns the
surrounding panel, heading/page count, toolbar and open/closed state.

After mounting Word interactions, or receiving a PDF `ready` event, return
`page_thumbnails_command(config, "page-list", enabled=panel_open)`. Calling it
again updates visibility and the Word font preference without replacing the
live document. Labels follow later renders. The list supports click, arrow
keys, Home and End, highlights the current page, and scrolls that highlight
into its own viewport without moving the host window.

Serve `public/page-thumbnails.js` at `./page-thumbnails.js`, or pass a
`runtime_url` to the command. Include `public/page-thumbnails.css` (already
imported by the standalone stylesheet). Its optional `--office-pages-text`,
`--office-pages-hover`, `--office-pages-accent`, `--office-pages-selected` and
`--office-pages-font-size` properties let a host supply its theme. Panel width,
placement and responsive breakpoints are host policy and are not measured by
the thumbnail component.

Word thumbnails clone the actual paginated SVG with isolated definition IDs
and the selected font family. PDF thumbnails use separate small PDF.js
canvases, with two concurrent jobs and up to 24 cached previews. Only nearby
pages are rendered. Rotating a PDF or changing Word font mode invalidates
previews. Disabling the list cancels pending work; replacing or disposing the
document retires listeners, observers, requests and canvases automatically.
Call `dispose_page_thumbnails_command(surface_id)` before independently
removing the list while keeping the document mounted. Excel has no page list.

## Materialization and resolution

After receiving selection JSON, choose an explicit context policy:

```mbt
let full = @viewer.materialize_selection_full(selection_json)
let preview = @viewer.materialize_selection_preview(selection_json, 1000)
let deferred = @viewer.materialize_selection_deferred(selection_json)
```

These return `office.selection.materialized.v1`. `full` retains exact text;
`preview` contains a deterministic head/tail excerpt with UTF-16 omission counts;
`deferred` omits selected text. For PDF, all three modes project only the
selected text and one complete PDF.js reference into the Agent payload. They
omit repeated document identity, per-item font/geometry metadata, tagged
semantics, and page details; the raw selection callback and resolver still
provide those details when an application needs them. Use `materialize_selection`
with `SelectionMaterializationOptions` to configure the excerpt and DOCX/XLSX
compact-anchor bounds. PDF materializations always omit per-item anchors. Hosts
must check `source.readable` before relying on a reference.

The upstream `pagelayout/svg` backend has no OOXML source map. The viewer adds
one only where its annotated-reader paragraph/run projection joins the rendered
text exactly. Otherwise Word selection retains its physical page/fragment
range with `stability: "physical-only"` and omits a canonical selector. The same
fail-closed join governs contextual comment placement. A proven body-paragraph
prefix can remain as comment context when a later unsupported block prevents
a complete join. This context carries no logical selectors, offsets or
actionability, and is never claimed to be the exact commented character range.
Running headers and footers are excluded from body-paragraph matching.

## Selecting comments and Notes

The same `interactions_command` callback now emits `kind: "annotation"` when the
user selects a contextual comment card or highlights part of its body. Keep this
distinct from `word` and `excel` document selections. The envelope identifies the
source file and SHA-256, annotation ID/type, author/thread metadata, original
source anchors, and the selected annotation-body UTF-16 span. It never reuses a
paragraph or cell selector as a comment-body address.

`select_annotation_command(surface_id, item, on_result~)` supports host-owned
annotation lists, including entries that cannot be positioned in the preview.
`annotation_selection(item, start_utf16~, end_utf16~, scope="text")` constructs
the same projection without the browser. Omitting offsets selects the whole
body. Browser commands verify that the item still belongs to the installed
snapshot; pure projection assumes a host-provided item from that snapshot.

Full/preview/deferred materialization all accept this selection kind. Preview
and deferred retain metadata and the body span without carrying the unselected
comment body. Resolving a deferred annotation uses `read_annotations` on the
same file bytes/hash and item ID; `source.readable` remains false for the existing
`office.selector/1` document readback route. See the
[annotation selection contract](selection-model.md#annotation-body-selections).

DOCX/XLSX selections expose the existing file and `office.mbt` selectors. A
materialized PDF selection keeps only `selection.schema`, `selection.format`,
`selection.kind`, `selection.text`, and `source.readable` plus
`source.reference` using `office.pdf.text-range.v1`; it has no `structure`
block or `office.mbt` selector fields. Resolve it with:

```mbt
let result = @viewer.resolve_pdf_selection(surface_id, materialized_json)
```

This is an asynchronous `Result[String, String]` API returning a reconstructed
full `office.selection.v1` envelope. It accepts full, preview, or deferred
materializations, a raw selection, or a standalone PDF reference. It validates
the path, SHA-256, byte length, page count, parser/version, extraction options,
and endpoints, then reads the PDF.js text items again. It does not depend on
retaining selected text in an application cache.

Resolution requires an active PDF surface. For a persisted reference, reopen
identical bytes with the same file path and PDF.js version, then resolve it
against that surface. A changed document or parser version requires a new
selection. The detailed anchors may be bounded to 64 entries while the
complete start/end reference continues to cover the full selection.

## Assets and lifecycle

Run `npm ci` and `moon run tools/prepare-browser-assets.mbtx` before building
the standalone app. An embedded host must serve `public/pdf-runtime.js`,
`public/vendor/pdfjs/`, and the PDF.js viewer stylesheet. The default runtime
URL is `./pdf-runtime.js`, relative to `document.baseURI`. Pass `runtime_url`
to `load_pdf_command` when the host uses a different asset location. The
runtime resolves its worker, CMaps, standard fonts, WASM codecs, and ICC profiles
relative to its own URL, making a nested deployment path possible without a CDN.

The PDF.js `.page` element requires `box-sizing: content-box`; protect that rule
from global CSS resets. `.pdf-container` must be positioned absolutely inside a
bounded surface and own its scrolling. The demo stylesheet contains these rules.

PDF.js owns the children of the PDF surface. Preserve `surface_html` while
updating toolbar/selection state so Rabbita does not replace that subtree.
Mounting another document retires the previous runtime. Call
`dispose_interactions_command(surface_id)` while the old root still exists
before removing or replacing it. Cleanup cancels the worker/loading task,
listeners, selection timer, and page cache. Multiple surfaces use separate
sessions and unique ids.

PDF document JavaScript, XFA, and editing are disabled. PDF annotations use
the read-only display layer; external links open in a separate tab. OCR and
inferred table/paragraph extraction are outside the current PDF surface.

## Fit to the available width

Call `fit_width_command(surface_id, kind, on_zoom)` after mounting a DOCX or XLSX
surface. It measures the page or active sheet and the available host width,
including padding and the current comment gutter. Resize observers keep the
page fitted when its panel or gutter changes size. The callback reports the
integer percentage for the host toolbar and state; preserve that percentage
in `SurfaceConfig.zoom`. Fitting caps enlargement at 100%.

After explicit zoom or a pinch gesture, call the same command with
`enabled=false` to retain the user's scale. Re-enable it when Fit width is
chosen. Disposal removes pending frames and observers. DOCX defaults to fitting
in the standalone viewer; XLSX stays at 100% unless the user requests fitting.
PDF retains its PDF.js `pdf_fit_width_command` and scale state events.
