///|
pub(all) enum DocumentKind {
  Word
  Excel
  Pdf
} derive(Eq)

///|
pub extend DocumentKind with Eq::{equal, not_equal}

///|
/// The renderer output consumed by both the standalone shell and OpenSeek.
/// `file` is owned by the host. The standalone demo uses a mock path; an
/// embedded host should pass the real path used by `office.mbt`.
pub(all) struct RenderedDocument {
  kind : DocumentKind
  name : String
  file : String
  surface_html : String
  item_count : Int32
  warning : String
  annotations : Json
} derive(Eq)

///|
pub extend RenderedDocument with Eq::{equal, not_equal}

///|
pub fn document_kind(name : String) -> DocumentKind? {
  let lower = name.to_lower()
  if lower.has_suffix(".docx") {
    Some(Word)
  } else if lower.has_suffix(".xlsx") {
    Some(Excel)
  } else if lower.has_suffix(".pdf") {
    Some(Pdf)
  } else {
    None
  }
}

///|
/// Extract the body while preserving the renderer's generated style block.
/// `xlsx2html` returns a complete HTML document; a host surface only needs its
/// body fragment because the host owns the page shell and stylesheet.
fn extract_html_body(document : String) -> String {
  let style = match (document.find("<style>"), document.find("</style>")) {
    (Some(start), Some(end)) => document[start:end + 8].to_owned()
    _ => ""
  }
  match (document.find("<body>"), document.find("</body>")) {
    (Some(start), Some(end)) => style + document[start + 6:end].to_owned()
    _ => style + document
  }
}

///|
/// Keep formulas alongside the rendered result so the browser formula bar can
/// show the source expression without changing the HTML emitted by
/// `mbtexcel/xlsx2html`.
fn excel_freeze_metadata(workbook : @xlsx.Workbook) -> String raise {
  let sheets : Map[String, Json] = Map([])
  for name in workbook.get_sheet_list() {
    let panes = workbook.get_panes(name)
    sheets[name] = Json::object({
      "rows": json_number(
        if panes.freeze && panes.y_split > 0 {
          panes.y_split
        } else {
          0
        },
      ),
      "columns": json_number(
        if panes.freeze && panes.x_split > 0 {
          panes.x_split
        } else {
          0
        },
      ),
    })
  }
  Json::object({ "sheets": Json::object(sheets) })
  .stringify()
  .split("<")
  .join("\\u003c")
}

///|
fn excel_formula_metadata(
  workbook : @xlsx.Workbook,
  max_rows : Int32,
  max_cols : Int32,
) -> String raise {
  let sheets : Array[(String, Json)] = []
  for name in workbook.get_sheet_list() {
    let formulas : Map[String, Json] = Map([])
    match workbook.sheet(name) {
      Some(worksheet) => {
        let max_row = if worksheet.max_row() > max_rows {
          max_rows
        } else {
          worksheet.max_row()
        }
        let max_col = if worksheet.max_col() > max_cols {
          max_cols
        } else {
          worksheet.max_col()
        }
        if max_row > 0 && max_col > 0 {
          for row in 1..<=max_row {
            for col in 1..<=max_col {
              let cell = @mbtexcel.coordinates_to_cell_name(col, row)
              match workbook.get_cell_formula(name, cell) {
                Some(formula) => formulas[cell] = Json::string(formula)
                None => ()
              }
            }
          }
        }
      }
      None => ()
    }
    if formulas.length() > 0 {
      sheets.push((name, Json::object(formulas)))
    }
  }
  // A formula is workbook input. Escaping `<` keeps a hostile `</script>`
  // sequence inside the JSON data block instead of ending the script element.
  Json::object({ "sheets": Json::object(Map(sheets)) })
  .stringify()
  .split("<")
  .join("\\u003c")
}

///|
fn excel_parse_double_opt(text : String) -> Double? {
  let number = @string.parse_double(text) catch { _ => return None }
  Some(number)
}

///|
/// Convert formula results that arrived through the string-only formula
/// writer API back into typed numeric cells for the preview workbook. The
/// reader keeps the original workbook untouched by callers because the
/// caller can produce this workbook solely for rendering. Keeping the
/// metadata pass before this function means the browser still receives the
/// formula source after `set_cell_value` removes the formula from the visual
/// copy.
fn normalize_formula_values_for_preview(
  workbook : @xlsx.Workbook,
) -> Unit raise {
  for name in workbook.get_sheet_list() {
    match workbook.sheet(name) {
      Some(worksheet) =>
        for reference in worksheet.formula_refs() {
          match workbook.calc_cell_typed(name, reference) {
            Some(@xlsx.Numeric(number)) =>
              worksheet.set_cell_value(reference, @xlsx.Numeric(number))
            Some(@xlsx.String(text)) =>
              // A few producers serialize a numeric cached result as a text
              // value. Only coerce an unambiguous finite number; formulas
              // whose semantic result is text remain text.
              match excel_parse_double_opt(text) {
                Some(number) if !number.is_nan() && !number.is_inf() =>
                  worksheet.set_cell_value(reference, @xlsx.Numeric(number))
                _ => ()
              }
            _ => ()
          }
        }
      None => ()
    }
  }
}

///|
/// Convert the OOXML EMU size used by drawing anchors into CSS pixels. Excel
/// uses 9,525 EMU per CSS pixel for worksheet drawings.
fn emu_to_px(value : Int32) -> Int32 {
  if value <= 0 {
    1
  } else {
    let pixels = value.to_double() / 9525.0
    let rounded = (pixels + 0.5).to_int()
    if rounded < 1 {
      1
    } else {
      rounded
    }
  }
}

///|
fn excel_json_int(value : Int32) -> Json {
  Json::number(value.to_double(), repr=value.to_string())
}

///|
/// Describe sheet drawings for the browser-side anchor layout. The object
/// order follows xlsx2html's output: pictures first, then charts.
fn excel_object_metadata(workbook : @xlsx.Workbook) -> String raise {
  let sheets : Array[(String, Json)] = []
  for name in workbook.get_sheet_list() {
    let objects : Array[Json] = []
    match workbook.sheet(name) {
      Some(worksheet) => {
        for cell in workbook.get_picture_cells(name) {
          for image in workbook.get_pictures(name, cell) {
            objects.push(
              Json::object({
                "kind": Json::string("image"),
                "anchor": Json::string(cell),
                "offset_x": excel_json_int(image.offset_x),
                "offset_y": excel_json_int(image.offset_y),
                "width_emu": excel_json_int(image.width_emu),
                "height_emu": excel_json_int(image.height_emu),
                "width_px": excel_json_int(emu_to_px(image.width_emu)),
                "height_px": excel_json_int(emu_to_px(image.height_emu)),
              }),
            )
          }
        }
        for chart in worksheet.charts() {
          objects.push(
            Json::object({
              "kind": Json::string("chart"),
              "anchor": Json::string(chart.reference),
              "offset_x": excel_json_int(chart.offset_x),
              "offset_y": excel_json_int(chart.offset_y),
              "width_emu": excel_json_int(chart.width_emu),
              "height_emu": excel_json_int(chart.height_emu),
              "width_px": excel_json_int(emu_to_px(chart.width_emu)),
              "height_px": excel_json_int(emu_to_px(chart.height_emu)),
            }),
          )
        }
      }
      None => ()
    }
    if objects.length() > 0 {
      sheets.push((name, Json::array(objects)))
    }
  }
  Json::object({ "sheets": Json::object(Map(sheets)) })
  .stringify()
  .split("<")
  .join("\\u003c")
}

///|
/// Replace one upstream unsupported-chart placeholder. Keeping this small
/// post-processing step lets ordinary charts continue to use xlsx2html while
/// the viewer supplies the extra composite-chart behavior it needs.
fn replace_chart_placeholder(
  document : String,
  anchor : String,
  replacement : String,
) -> String {
  let marker = "<figure class=\"chart-placeholder\" data-anchor=\"" +
    @xlsx2html.escape_html_attribute(anchor) +
    "\">"
  match document.find(marker) {
    Some(start) =>
      match document[start:].find("</figure>") {
        Some(end_offset) => {
          let end = start + end_offset + "</figure>".length()
          document[0:start].to_owned() + replacement + document[end:].to_owned()
        }
        None => document
      }
    None => document
  }
}

///|
fn render_excel_document(
  workbook : @xlsx.Workbook,
  title : String,
  max_rows : Int32,
  max_cols : Int32,
) -> String raise {
  let mut rendered = @xlsx2html.render_workbook(
    workbook,
    title~,
    max_rows~,
    max_cols~,
    include_images=true,
    calc=true,
  )
  for worksheet in workbook.sheets() {
    for chart in worksheet.charts() {
      let kinds = @inspect.chart_kinds(chart.xml)
      // xlsx2html already draws single bar/line/pie charts. Composite plots
      // and area plots need the viewer renderer so they are not left as a
      // labeled placeholder.
      if kinds.length() > 1 || kinds.contains("areaChart") {
        match viewer_chart_figure(workbook, chart.xml, chart.reference) {
          Some(figure) =>
            rendered = replace_chart_placeholder(
              rendered,
              chart.reference,
              figure,
            )
          None => ()
        }
      }
    }
  }
  rendered
}

///|
/// Keep page boundaries as DOM elements so native browser text selection can
/// report the physical page containing each selected fragment.
fn word_surface(pages : Array[String]) -> String {
  let output = StringBuilder()
  output.write_string("<div class=\"word-pages\">")
  for index, page in pages {
    output.write_string(
      "<article class=\"word-page\" data-page=\"\{index + 1}\">",
    )
    output.write_string(page)
    output.write_string("</article>")
  }
  output.write_string("</div>")
  output.to_string()
}

///|
/// Render one Office package with the same layout engines used by the
/// standalone demo. This function has no browser or OpenSeek state, so an
/// embedded JS frontend can call it after its native host supplies bytes.
pub fn render_document(
  name : String,
  file : String,
  bytes : Bytes,
  max_rows~ : Int32 = 1000,
  max_cols~ : Int32 = 256,
) -> Result[RenderedDocument, String] {
  if max_rows < 1 || max_cols < 1 || max_rows > 262144 / max_cols {
    return Err(
      "Preview bounds must be positive and contain at most 262144 cells.",
    )
  }
  let kind = match document_kind(name) {
    Some(kind) => kind
    None => return Err("Only .docx, .xlsx, and .pdf files are supported.")
  }
  try {
    match kind {
      Word => {
        let document = @docx.read_docx(bytes)
        let model = @paginate.paginate(
          document.blocks,
          section=document.section,
          furniture=document.furniture,
        )
        let rendered_pages = @svg.render_pages(model)
        let annotated = Some(@docx2html.read_docx_annotated(bytes)) catch {
          _ => None
        }
        let annotations = match annotated {
          Some(value) => word_annotations(name, file, value)
          None =>
            annotation_envelope(
              name,
              file,
              "docx",
              [],
              Json::object({ "completeness": Json::string("unavailable") }),
              [
                "The annotated Word reader failed. Comments are unavailable; this is not evidence that the document has no comments.",
              ],
            )
        }
        let index = match annotated {
          Some(value) => build_word_source_index_annotated(value)
          None => None
        }
        let annotations = bind_annotation_snapshot(annotations, bytes)
        let pages = match index {
          Some(index) => {
            let furniture = document.furniture
            // Running headers and footers are appended after body items.
            // Source paragraph matching must only consume the body model.
            let body_model = if furniture.header.is_empty() &&
              furniture.footer.is_empty() &&
              furniture.first_header.is_empty() &&
              furniture.first_footer.is_empty() &&
              furniture.even_header.is_empty() &&
              furniture.even_footer.is_empty() {
              model
            } else {
              @paginate.paginate(document.blocks, section=document.section)
            }
            decorate_word_pages(rendered_pages, body_model, index)
          }
          None => rendered_pages
        }
        Ok({
          kind,
          name,
          file,
          surface_html: render_warning_html(document.unsupported) +
          word_surface(pages) +
          annotations_script(annotations),
          annotations,
          item_count: pages.length(),
          warning: if document.unsupported.is_empty() {
            ""
          } else {
            "Preview limitation: " +
            document.unsupported[0] +
            (if document.unsupported.length() > 1 {
              " (additional limitations listed in the preview)"
            } else {
              ""
            })
          },
        })
      }
      Excel => {
        let workbook = @mbtexcel.read(bytes.exact_view())
        let annotations = bind_annotation_snapshot(
          excel_annotations(name, file, workbook, max_rows~, max_cols~),
          bytes,
        )
        let formula_metadata = excel_formula_metadata(
          workbook, max_rows, max_cols,
        )
        let freeze_metadata = excel_freeze_metadata(workbook)
        let object_metadata = excel_object_metadata(workbook)
        normalize_formula_values_for_preview(workbook)
        let rendered = render_excel_document(workbook, name, max_rows, max_cols)
        Ok({
          kind,
          name,
          file,
          surface_html: "<div class=\"sheet-document\">" +
          extract_html_body(rendered) +
          "<script type=\"application/json\" class=\"office-formula-metadata\">" +
          formula_metadata +
          "</script>" +
          "<script type=\"application/json\" class=\"office-freeze-metadata\">" +
          freeze_metadata +
          "</script>" +
          "<script type=\"application/json\" class=\"office-object-metadata\">" +
          object_metadata +
          "</script>" +
          annotations_script(annotations) +
          "</div>",
          annotations,
          item_count: workbook.get_sheet_list().length(),
          warning: "Excel cells are read-only. Formula results are normalized in the rendering copy when office.mbt can evaluate them; drawing anchors retain their source cell, offsets, and dimensions for browser layout. Cell addresses are derived from the rendered table coordinates and use the canonical office.xlsx selector shape.",
        })
      }
      Pdf =>
        Err(
          "PDF rendering is asynchronous; mount pdf_surface_html() and use load_pdf_command().",
        )
    }
  } catch {
    error => Err(error.to_string())
  }
}
