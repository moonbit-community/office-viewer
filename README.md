# Office Viewer

A read-only document viewer built with MoonBit and Rabbita. Office Viewer
provides standalone DOCX, XLSX, and PDF previews and a reusable browser surface
with structured selections for host applications and agents.

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

The workspace uses pinned sibling checkouts. With this repository checked out as
`office-viewer`, prepare the dependencies from its parent directory:

```sh
git clone https://github.com/moonbitlang/office.mbt.git
git -C office.mbt checkout 59596ed5181382494ae5e5a151e2e954a8350001
git clone https://github.com/moonbit-community/rabbita.git
git -C rabbita checkout d31819b4b59653550af5fcd3c1bc5c424caae64a
```

Use the [official MoonBit toolchain](https://www.moonbitlang.com/download)
`0.10.14+7d59c7ec9` (moon `0.1.20260920`) and Node.js 24. From `office-viewer`,
resolve dependencies in this order; the Office workspace supplies the local XML
package referenced by `moon.work`:

```sh
moon -C ../office.mbt update
moon -C ../office.mbt check --target js ooxml/xml
moon -C ../rabbita update
moon update
npm ci --ignore-scripts
moon run tools/prepare-browser-assets.mbtx
moon check --target js
moon test --target js -p moonbit-community/office-viewer/viewer
moon test --target js -p moonbit-community/office-viewer/cmd/browser
```

Start the development server from `office-viewer`:

```sh
moon -C ../rabbita run --target native warren -- \
  -C "$PWD" dev --direct --browser-entry cmd/browser --server-entry "" --port 4311
```

Open <http://127.0.0.1:4311/>. To build the same release artifact as CI, install
Terser in a temporary directory and put it on `PATH`:

```sh
minifier="$(mktemp -d)"
npm install --prefix "$minifier" --ignore-scripts --no-audit --no-fund terser@5.51.2
PATH="$minifier/node_modules/.bin:$PATH" bash scripts/build-pages.sh
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
The workflow pins the two sibling revisions, the official compiler/core archives
and their checksums, Terser, and npm browser dependencies. Update these pins together after compatibility
checks. The deployed `build-info.json` identifies the exact source commit and
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
`office.mbt` selectors. References record the file path, SHA-256, parser
version, text-item endpoints, and page geometry. Hosts can resolve them against
the same file bytes using `resolve_pdf_selection`. The demo's Selection panel
exposes Full, Preview, and Deferred output and PDF reference resolution. See
[PDF preview](docs/pdf-preview.md) for the PDF contract and browser tests.

Existing DOCX comments/w15 replies and Excel traditional Notes are available
in a default-visible document gutter and the optional Comments overview, and through
`@viewer.read_annotations` or
`RenderedDocument.annotations`. The `office.annotations.v1` contract preserves
missing metadata and explicit support limits. Word navigation shows only proven
paragraph context; Notes beyond preview bounds remain in the list and JSON.
Comment and reply bodies appear on file load without opening Comments or
clicking a marker. Select a whole comment or highlight any body text to capture an
`office.selection.v1` annotation-body selection, including source identity and
UTF-16 offsets. Preview/deferred modes retain the reference without leaking the
unselected body.
See [`docs/annotations-model.md`](docs/annotations-model.md) for the API,
unsupported modern-comment formats, and compatible sibling checkout versions.

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
