# Rendering support

The viewer uses published docx2html 0.7.0, mbtexcel 0.3.0, and pagelayout 0.8.0.
This change adds browser freeze handling, configurable Excel preview bounds,
and visible DOCX diagnostics using those published APIs. It does not require
unreleased upstream HTML attributes: frozen pane settings come from the parsed
workbook in a separate inert JSON block.

The accompanying office.mbt changes improve linear gradient fills, borders,
theme/indexed colors, and DOCX body inline pictures. They must be merged and
released upstream, followed by a viewer dependency update, before a registry
build contains those engine fixes. A local MoonBit workspace can exercise the
viewer with the upstream source modules without changing release pins.

## Known limits

- DOCX floating drawing anchors and text wrapping, nested tables, paragraph
  section geometry, footnotes/endnotes, equations, and text boxes need additional
  layout support. Diagnostics report detected omissions; they do not render them.
- Upstream body inline pictures participate in text wrapping, baseline placement,
  line height and pagination. Table-cell and DOCX header/footer picture reading
  remain follow-up work.
- Linear Excel gradients with two colors are supported by the upstream patch.
  Path gradients and full conditional formatting remain incomplete. CSS preserves
  border weights and approximates dash-dot variants as dashed lines.
- The viewer's frozen rows/columns operate within each worksheet scroll area.
  Excel's separately scrollable split-pane UI is not implemented.
- Preview bounds limit the displayed table, formulas and Notes navigation; they
  do not stream or virtualize an arbitrarily large workbook. Cropped tables retain
  the renderer's visible truncation message. Notes outside the window remain in
  annotation data.

## Verification

Run `moon test --target js`, `npm run test:browser`, and `npm run build:pages`.
Browser tests use Playwright Chromium; `CHROME_EXECUTABLE` can select an installed
Chrome, and `PLAYWRIGHT_MODULE` can select an externally supplied Playwright module.
