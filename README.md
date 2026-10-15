# Office Viewer

A read-only document viewer built with MoonBit and Rabbita. Office Viewer
provides standalone DOCX, XLSX, and PDF previews and a reusable browser surface
with structured selections for host applications and agents.

DOCX/XLSX interaction state, UI construction, source mapping, and selection
logic live in MoonBit, using Rabbita's typed DOM interfaces and small browser
bindings. See the [browser runtime](docs/browser-runtime.md) for the boundaries
and verification workflow.

[Live demo](https://moonbit-community.github.io/office-viewer/)

## Supported formats

| Format | Rendering | Selection |
| --- | --- | --- |
| DOCX | `office.mbt` PageModel and paginated SVG | Native text selection with validated paragraph/run references where available |
| XLSX | `office.mbt` workbook reader and HTML renderer | Cells and rectangular ranges with worksheet names and canonical selectors |
| PDF | PDF.js canvas, text, and annotation layers | Native text selection with page/item ranges, document identity, geometry, and available tagged structure |

The viewer supports zoom and search. PDF previews also provide page navigation,
fit to width, rotation, and password prompts. Files are read locally in the
browser; the standalone demo uses a mock file path in its selection metadata.
Excel previews preserve workbook styles, merged cells, images, charts, and
formula results. Supported uncached formulas are evaluated by `office.mbt`;
unsupported formulas remain visibly marked as pending.

## Run locally

Use the [official MoonBit toolchain](https://www.moonbitlang.com/download)
`0.10.15+943464bfa` and Node.js 24. Dependencies resolve from the Mooncakes
registry; sibling checkouts are not required.

```sh
moon update
npm ci --ignore-scripts
moon run tools/prepare-browser-assets.mbtx
moon check --target js
moon test --target js
npx --no-install playwright install chromium
npm run test:browser
```

Start the development server:

```sh
moon install moonbit-community/warren@0.4.5
warren \
  -C "$PWD" dev --direct --browser-entry cmd/browser --server-entry "" --port 4311
```

Open <http://127.0.0.1:4311/>. With Warren installed as above, build the same release artifact as CI with:

```sh
npm run build:pages
```

The generated `dist/` is ignored by Git; an optional first argument selects a
different output directory inside the repository (required by Warren). The script
uses Warren to compile/minify the browser entry and copy `public/`, including the existing bundled fonts. It fingerprints
JS/CSS filenames and writes `build-info.json` with the source revision and full
asset SHA-256 hashes. It does not regenerate fonts. Run
`moon run tools/prepare-browser-assets.mbtx` after updating npm dependencies;
the generated PDF.js worker, CMaps, fonts, codecs, and icon assets are copied
from `public/vendor/` into the release artifact.

## Continuous deployment

[The Pages workflow](.github/workflows/pages.yml) checks, tests, and builds on
pull requests and every push to `main`. PRs never deploy. After a successful main
build (or a manual run on `main`), it uploads the generated temporary artifact
and deploys it with GitHub Pages. Build output is not committed to the repository.
The workflow pins the published Mooncakes dependencies, official compiler/core
archives and their checksums, Warren (with its bundled minifier), and npm browser dependencies.
Update these pins together after compatibility checks. The deployed `build-info.json` identifies the exact source commit and
asset contents. PDF.js and icon assets are generated before the source build;
they are served locally from the deployed directory, including under subpaths.

## Structured selections and annotations

The current selection layer is deliberately read-only: Word uses browser text
selection inside positioned page SVG, and Excel uses rectangle selection over
cells. Word's source metadata is built in the viewer by reading the same bytes
through `office.mbt`'s annotated DOCX reader and joining its paragraph/run
projection with the PageModel text. The join is fail-closed, so an uncertain
paragraph remains physical-only instead of receiving a guessed selector. Both
formats produce the versioned `office.selection.v1` JSON envelope shown in
[`docs/selection-model.md`](docs/selection-model.md); the inspector displays
that envelope and can copy it through the browser clipboard for the later
host application's Agent conversation.

Hosts can choose `@viewer.materialize_selection_full`,
`@viewer.materialize_selection_preview`, or
`@viewer.materialize_selection_deferred` to control how much selected content
enters an Agent context. The APIs preserve the existing `office.mbt` file and
selector fields and make UTF-16 omission counts explicit; see
[`docs/embedding.md`](docs/embedding.md).

PDF selections use `office.pdf.text-range.v1` references rather than
`office.mbt` selectors. Agent materializations keep the selected text and a
single reference with the file path, SHA-256, parser version, and text-item
endpoints. Hosts can resolve that reference against the same file bytes using
`resolve_pdf_selection`; the raw selection and resolver result retain detailed
geometry and tagged metadata. The demo's Selection panel exposes Full, Preview,
and Deferred output and PDF reference resolution. See
[PDF preview](docs/pdf-preview.md) for the PDF contract and browser tests.

Existing DOCX comments/w15 replies and Excel traditional Notes are available
in a document gutter and the optional Comments overview, and through
`@viewer.read_annotations` or
`RenderedDocument.annotations`. The `office.annotations.v1` contract preserves
missing metadata and explicit support limits. Word navigation shows only proven
paragraph context; Notes beyond preview bounds remain in the list and JSON.
The gutter opens alongside the document when space permits; on narrow layouts,
open its heading or a source marker to read comments and replies. Quote a whole
comment or highlight any body text to capture an
`office.selection.v1` annotation-body selection, including source identity and
UTF-16 offsets. Preview/deferred modes retain the reference without leaking the
unselected body.
See [`docs/annotations-model.md`](docs/annotations-model.md) for the API,
unsupported modern-comment formats, and tested dependency versions.

The reusable surface lives in [`viewer/`](viewer/). Its renderer and Rabbita
surface are host-neutral: a standalone host can provide browser file bytes,
while OpenSeek can provide bytes and the real `office.mbt` file path from its
native backend. The intended OpenSeek integration mounts this same JS/Rabbita
surface inside the desktop frontend; it does not open a second webpage or
iframe.

Word surfaces include a visible limitations banner and inert
`office.render.diagnostics.v1` JSON when the layout reader reports unsupported
content. Warning coverage depends on the installed `pagelayout` release; an
empty warning list does not certify complete OOXML rendering. Floating drawings,
nested tables, and paragraph-level section geometry still need engine work.

Excel freezes authored rows/columns in a worksheet scroll area, including after
zoom and worksheet switches. Split panes are not treated as frozen panes.
Embedding hosts can set `render_document(..., max_rows=..., max_cols=...)`;
the defaults are 1000 rows and 256 columns with a 262144-cell per-sheet ceiling.
The table, formula metadata, and Notes preview flags use the same bounds.
Notes outside the preview remain available in the annotation list.
See [rendering support](docs/rendering-support.md) for version boundaries and
known limitations.

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

From the Office checkout, regenerate bundled fonts only after changing the
upstream font data:

```sh
moon run --target native ./pagelayout/cmd/pagelayout -- \
  export-fonts ../office-viewer/public/fonts
```

PDF fonts, CMaps, and images are handled by PDF.js. Embedded fonts come from
the opened document; standard fonts and CMaps are included with the PDF.js
assets and their upstream license notices. Image-only PDFs display normally;
OCR and inferred paragraph/table extraction are not included.
