# Changelog

## 0.2.0 (unreleased)

- Freeze authored Excel rows and columns while scrolling, including after zoom
  and worksheet switches.
- Allow embedding hosts to configure the Excel preview window, with a per-sheet
  cell budget shared by tables, formula metadata and Notes preview flags.
- Show DOCX layout limitations inside the reusable surface with counted,
  machine-readable diagnostics.
- Escape every `<` in inert formula, object, annotation and diagnostics JSON.
- Resolve current published Office modules and pin the compatible compiler and
  Warren builder; build and test without sibling checkouts.
