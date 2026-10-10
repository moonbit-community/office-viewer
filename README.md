# Office Viewer

Standalone, read-only Word and Excel preview surface for the OpenSeek
integration work.

Demo: <https://moonbit-community.github.io/office-viewer/>

The browser entry is `cmd/browser`. DOCX files are parsed by
`moonbitlang/pagelayout/docx`, paginated into a `PageModel`, and emitted as
page SVG. XLSX files are parsed by `moonbitlang/mbtexcel` and rendered by
`xlsx2html`, preserving workbook styles, merged cells, images, charts, and
formula results. Cached OOXML values are used when present; uncached formulas
are evaluated by `office.mbt` where the formula engine supports them, and an
unsupported formula remains visibly marked as pending instead of being
silently replaced with an incorrect value.

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
`0.10.14+7d59c7ec9` (moon `0.1.20260920`) and Node.js 22. From `office-viewer`,
resolve dependencies in this order; the Office workspace supplies the local XML
package referenced by `moon.work`:

```sh
moon -C ../office.mbt update
moon -C ../office.mbt check --target js ooxml/xml
moon -C ../rabbita update
moon update
moon check --target js
moon test --target js -p moonbit-community/office-viewer/viewer
moon test --target js -p moonbit-community/office-viewer/cmd/browser
```

Start the development server from `office-viewer`:

```sh
moon -C ../rabbita run --target native warren -- \
  -C "$PWD" dev --direct --browser-entry cmd/browser --server-entry "" --port 4310
```

Open <http://127.0.0.1:4310/>. To build the same release artifact as CI, install
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
asset SHA-256 hashes. It does not regenerate fonts.

## Continuous deployment

[The Pages workflow](.github/workflows/pages.yml) checks, tests, and builds on
pull requests and every push to `main`. PRs never deploy. After a successful main
build (or a manual run on `main`), it uploads the generated temporary artifact
and deploys it with GitHub Pages. Build output is not committed to the repository.
The workflow pins the two sibling revisions, the official compiler/core archives
and their checksums, and Terser. Update these pins together after compatibility
checks. The deployed `build-info.json` identifies the exact source commit and
asset contents.

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

Hosts can choose `@viewer.materialize_selection_full`,
`@viewer.materialize_selection_preview`, or
`@viewer.materialize_selection_deferred` to control how much selected content
enters an Agent context. The APIs preserve the existing `office.mbt` file and
selector fields and make UTF-16 omission counts explicit; see
[`docs/embedding.md`](docs/embedding.md).

Existing DOCX comments/w15 replies and Excel traditional Notes are available
in contextual cards and the read-only Comments panel, and through
`@viewer.read_annotations` or
`RenderedDocument.annotations`. The `office.annotations.v1` contract preserves
missing metadata and explicit support limits. Word navigation shows only proven
paragraph context; Notes beyond preview bounds remain in the list and JSON.
Select a whole comment or highlight part of a contextual card to capture an
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
