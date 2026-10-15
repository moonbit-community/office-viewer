# Changelog

## 0.2.0 (unreleased)

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
