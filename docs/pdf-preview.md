# PDF preview

Office Viewer uses PDF.js 6.4.299 for PDF rendering and text extraction. MoonBit
owns the reusable Rabbita commands, document dispatch, materialization API,
and standalone application state. The browser runtime connects PDF.js text
items to native browser selections.

## Rendering and selection

The viewer renders paginated canvases, text layers, and read-only annotation
layers. It supports page navigation, rotation, zoom, fit to width, text search,
and password prompts. Font programs embedded in a document are rendered by
PDF.js; CMaps, standard fonts, image codecs, and ICC profiles are served locally.

Selection mapping uses the PDF.js text layer's item sequence and compares it
with the extracted text sequence. A mismatch does not receive a guessed
reference. Item indices refer to the original `getTextContent().items` array,
including marked-content entries, and offsets are JavaScript UTF-16 code units.
The runtime pins `includeMarkedContent: true` and `disableNormalization: true`.

The selection carries an `office.pdf.text-range.v1` reference. Its file path,
SHA-256, byte length, page count, PDF.js version, and extraction options identify
the document and parser context. An item index is not a stable PDF object id
or an `office.mbt` selector. References must be recreated when document bytes
or the parser version change.

Canonical text preserves selected item strings. `pdfjs-items-eol-v1` inserts a
newline after an item with `hasEOL`, except the final referenced item, and two
newlines between pages. `browser_text` separately records the browser's selected
string. Text-item order is not a guarantee of human reading order.

Geometry uses PDF page user coordinates, independent of viewer zoom, rotation,
scrolling, and CropBox placement. The page's `UserUnit` is recorded explicitly.
Captured DOM ranges have precision `selected-dom-range`; reference resolution
without a DOM range returns `whole-text-item` metric bounds. The latter covers
the entire text item even when only part of its text is selected, and is not a
per-character or glyph-ink outline. Roles, language, and alternate text are
carried only when they are supplied by the PDF structure tree.

No OCR or inferred paragraph/table model is provided. Image-only pages display
normally but have no selectable text. PDF document JavaScript, XFA, and editing
are disabled. The PDF.js rendering engine is reused directly; this feature
does not claim to implement a new PDF renderer in MoonBit.

## Browser regression harness

Start the demo or serve a static build, then run:

```sh
moon run tools/pdf-browser-test.mbtx \
  http://127.0.0.1:4311 /absolute/path/to/chromium
```

The harness launches headless Chromium through CDP and writes reports and
screenshots under `_build/pdf-browser-<timestamp>/`. It creates an original PDF
fixture containing tagged text, over 64 text anchors, a link annotation,
multiple pages, a rotated CropBox with `UserUnit`, and an image-only page.

The checks exercise painted canvas pixels, mouse text selection, DOM cursor
extension, full/preview/deferred output, cache-independent resolution, metadata,
outline, search, coordinate stability under zoom/rotation, mobile layout,
malformed files, and loading-session replacement. Cursor extension uses the
browser's `Selection.extend` API; it does not validate physical Shift/Arrow
keyboard behavior or OS caret-browsing settings.

Optional PDF.js upstream fixtures can be tested without copying them into this
repository:

```sh
moon run tools/pdf-browser-test.mbtx \
  http://127.0.0.1:4311 /absolute/path/to/chromium \
  /absolute/path/to/pdf.js/test/pdfs
```

The additional fixtures are `pr6531_1.pdf` (password retry, unlock, and cancel),
`90ms_rksj_h_sample.pdf` (Japanese text and predefined CMaps), `tracemonkey.pdf`
(embedded fonts in a 14-page paper), and `vertical.pdf` (vertical text and
whole-item geometry). The fixture directory is read-only. Password and text
expectations follow PDF.js's upstream API tests. These tests are a focused
integration regression suite, not a claim of exhaustive PDF compatibility or
cross-browser certification.

Public MoonBit contract tests in `viewer/pdf_selection_test.mbt` cover format
dispatch, the three materialization modes, complete references after anchor
truncation, malformed endpoints, identity validation, and selector separation.
