# Office Viewer

Standalone, read-only Word and Excel preview surface for the OpenSeek
integration work.

Demo: <https://moonbit-community.github.io/office-viewer/>

The browser entry is `cmd/browser`. DOCX files are parsed by
`moonbitlang/pagelayout/docx`, paginated into a `PageModel`, and emitted as
page SVG. XLSX files are parsed by `moonbitlang/mbtexcel` and rendered by
`xlsx2html`, preserving workbook styles, merged cells, images, charts, and
formula text within the renderer's current limits.

## Run locally

The MoonBit workspace currently uses the local `office.mbt` and `rabbita`
checkouts for the layout, reader, and browser runtime packages. Keep them next
to this repository as `../office.mbt` and `../rabbita` when running the commands
below.

From the Rabbita checkout:

```sh
moon run --target native warren -- \
  -C /workspace/OfficeViewer/office-viewer \
  dev --direct --browser-entry cmd/browser --server-entry "" --port 4310
```

Open <http://127.0.0.1:4310/>. A release bundle can be generated with:

```sh
moon run --target native warren -- \
  -C /workspace/OfficeViewer/office-viewer \
  build --browser-entry cmd/browser --server-entry "" --dist dist
```

The current selection layer is deliberately read-only: Word uses browser text
selection inside positioned page SVG, and Excel uses rectangle selection over
cells. Word's source metadata is built in the viewer by reading the same bytes
through `office.mbt`'s annotated DOCX reader and joining its paragraph/run
projection with the PageModel text. The join is fail-closed, so an uncertain
paragraph remains physical-only instead of receiving a guessed selector. Both
formats produce the versioned `office.selection.v1` JSON envelope shown in
[`docs/selection-model.md`](docs/selection-model.md); the inspector displays
that envelope and can copy it through the browser clipboard for the later
OpenSeek Agent conversation.

The reusable surface lives in [`viewer/`](viewer/). Its renderer and Rabbita
surface are host-neutral: a standalone host can provide browser file bytes,
while OpenSeek can provide bytes and the real `office.mbt` file path from its
native backend. The intended OpenSeek integration mounts this same JS/Rabbita
surface inside the desktop frontend; it does not open a second webpage or
iframe.

The layout engines still have documented OOXML gaps, including some nested
tables, content controls, tab-stop cases, and paragraph-level section breaks.
Those need to be surfaced as structured provenance before claiming complete
Word compatibility.

## Font policy

The document surface uses the same bundled font programs as the pagelayout
metrics. They are exported into `public/fonts` from `pagelayout/fontoutlines`
and loaded with `@font-face`, so a browser fallback cannot change the fixed SVG
character positions. The bundled Carlito, Liberation, and Noto Sans SC files
are distributed under the license notices in `public/fonts/LICENSES`.

System fonts remain an explicit local enhancement. The `Aa` button switches
the browser's glyph source to installed Calibri/Arial/Times New Roman/Courier
New or common CJK equivalents when available, while keeping the page model's
bundled coordinates and pagination unchanged. Missing local faces fall back to
the bundled files. The default mode never silently changes with the host
machine; the browser runtime records both bundled loading and the selected
font mode.

To regenerate the bundled font files after changing the upstream font data:

```sh
moon run --target native ./pagelayout/cmd/pagelayout -- \
  export-fonts /workspace/OfficeViewer/office-viewer/public/fonts
```
