# Changelog

## 0.2.3 (unreleased)

- Add a reusable Word/PDF page-thumbnail list with lazy rendering, bounded
  caching, scrolling synchronization, click and keyboard navigation, and
  cancellation on document replacement or disposal. Hosts retain their own
  toolbar, labels, container layout and open/closed state.

## 0.2.2

- Preserve Word selections until pointer release and keep comments attached to
  proven body paragraphs when tables or running headers/footers prevent a
  complete semantic source mapping.
- Refit PDF scale presets as the embedding panel resizes, retain manual zoom,
  and retire resize callbacks when closing a document.
- Delay PDF selection previews until pointer release, including releases
  outside the document, and discard cancelled or retired drags.
- Include browser-ready names, checksums and modification notices for bundled
  fonts while preserving all glyph, character-map and metric tables.
- Run font, PDF lifecycle and Word drag regressions in the browser CI command.

## 0.2.1

- Expose independent fit-width observation and retain manual zoom choices.
- Add fit-width and pinch-zoom integration to the standalone document controls.

## 0.2.0 (unreleased)

- Preserve the last valid Excel drag range when releasing over unrelated or
  malformed cells, and avoid rescanning the grid while moving within one cell.
- Move DOCX/XLSX browser interaction state, UI construction, source mapping,
  and frozen-pane calculations into MoonBit. Browser tests now execute the
  compiled production runtime, with the existing embedding API preserved.
- Keep author fallback avatars beside their names inside comment cards, and
  place the Word comment gutter 12 pixels from the page with no stray avatars.
- Collapse narrow comment gutters, with source markers opening the relevant
  thread and source excerpts keeping expanded cards connected to the document.
- Freeze authored Excel rows and columns while scrolling, including after zoom
  and worksheet switches.
- Allow embedding hosts to configure the Excel preview window, with a per-sheet
  cell budget shared by tables, formula metadata and Notes preview flags.
- Show DOCX layout limitations inside the reusable surface with counted,
  machine-readable diagnostics.
- Escape every `<` in inert formula, object, annotation and diagnostics JSON.
- Resolve current published Office modules and pin the compatible compiler and
  Warren builder; build and test without sibling checkouts.
