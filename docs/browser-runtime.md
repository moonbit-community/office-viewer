# Browser interaction runtime

The reusable DOCX/XLSX surface owns its interaction state in MoonBit, in
`internal/browser`. Standalone and embedded hosts use the same implementation
through the existing `viewer` commands; their public API and JSON contracts are
unchanged.

The runtime separates responsibilities by feature:

- `runtime.mbt` owns each mounted surface, event listeners, animation frames,
  callbacks, and disposal.
- `excel.mbt`, `excel_selection.mbt`, and `excel_freeze.mbt` build worksheet
  controls, map merged cells, place drawings, manage range selection, and
  calculate frozen row and column positions.
- `annotation_model.mbt` validates installed annotation identities and resolves
  reply threads and source context. The annotation view, avatar, layout, and
  styles files build and position the corresponding UI.
- `selection.mbt` converts native browser ranges into Word source anchors and
  annotation body spans. `visual.mbt` applies search, zoom, and font preferences.

DOM queries, node creation, event subscriptions, geometry, resize observation,
and animation frames use Rabbita's typed browser interfaces. `browser_ffi.mbt`
contains short bindings for browser facilities not covered by those interfaces:
native Selection/Range access, text/style properties, host callback properties,
font loading, and browser locale and grapheme APIs. Viewer rules and mutable
interaction state belong in MoonBit, rather than JavaScript strings.

Replacing a surface first retires its previous listeners and pending frame,
then installs the canonical rendered HTML and derives a fresh interaction
snapshot. Explicit disposal removes the listeners, observers, callback hooks,
and registry entry. Font-readiness callbacks check whether that snapshot is
still active before applying layout.

## Verification

`npm run test:browser` first compiles `tools/runtime-harness`, which exposes the
production MoonBit runtime to Playwright. Tests run those compiled entry points
in Chromium; they do not extract or evaluate implementation snippets from FFI
source. The harness is separate from the production browser entry.

Coverage includes merged-cell dragging, formula inspection, sheet switching,
native Word and annotation selections, source identity validation, author
avatars, comment gutter navigation, frozen panes, zoom, and disposal. Use
`moon test --target js` for the document and selection contract tests and
`npm run build:pages` to verify the production entry.
