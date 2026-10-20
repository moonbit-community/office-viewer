///|
/// The lazy surface carries stored cells, not a dense rectangle. The document
/// owns this inert data together with its HTML, so switching/closing a document
/// cannot leave a workbook in a global renderer cache.
fn virtual_excel_model(workbook : @xlsx.Workbook) -> Json raise {
  let formulas : Map[(String, String), String] = Map([])
  for sheet in workbook.sheets() {
    for reference in sheet.formula_refs() {
      if workbook.get_cell_formula(sheet.name(), reference) is Some(formula) {
        formulas[(sheet.name(), reference)] = formula
      }
    }
  }
  normalize_formula_values_for_preview(workbook)
  let styles : Map[String, Json] = Map([])
  let sheets : Map[String, Json] = Map([])
  for sheet in workbook.sheets() {
    let name = sheet.name()
    let (rows, columns) = virtual_sheet_bounds(sheet)
    let merges : Array[Json] = []
    for merge in workbook.get_merge_cells_info(name, without_values=true) {
      let (left, top) = @xlsx.cell_name_to_coordinates(merge.get_start_axis())
      let (right, bottom) = @xlsx.cell_name_to_coordinates(merge.get_end_axis())
      merges.push(Json([top, left, bottom, right]))
    }
    let cells : Array[Json] = []
    for cell in sheet.cells {
      let style_id = sheet.effective_style_id_rc(cell.row, cell.col)
      if !styles.contains(style_id.to_string()) {
        styles[style_id.to_string()] = Json::string(
          @xlsx2html.style_css(workbook.get_style(style_id)),
        )
      }
      let formula = formulas.get((name, cell.reference)).unwrap_or("")
      let value = workbook
        .get_cell_value_styled(name, cell.reference, style_id)
        .unwrap_or("")
      let (value, pending) = if value == "" && formula != "" {
        (workbook.calc_cell_value(name, cell.reference), false) catch {
          _ => ("=" + formula, true)
        }
      } else {
        (value, false)
      }
      cells.push(
        Json::array([
          json_number(cell.row),
          json_number(cell.col),
          Json::string(value),
          json_number(style_id),
          Json::string(formula),
          Json::boolean(pending),
        ]),
      )
    }
    let row_dimensions : Array[Json] = []
    for (row, dimension) in sheet.materialized_row_dimensions() {
      let id = dimension.style_id.unwrap_or(-1)
      styles[id.to_string()] = Json::string(
        @xlsx2html.style_css(workbook.get_style(Int32::max(0, id))),
      )
      row_dimensions.push(
        Json::array([
          json_number(row),
          json_number(
            dimension.height.map(@xlsx2html.row_height_px).unwrap_or(26),
          ),
          json_number(id),
        ]),
      )
    }
    let column_dimensions : Array[Json] = []
    for (column, dimension) in sheet.materialized_col_dimensions() {
      let id = dimension.style_id.unwrap_or(-1)
      styles[id.to_string()] = Json::string(
        @xlsx2html.style_css(workbook.get_style(Int32::max(0, id))),
      )
      column_dimensions.push(
        Json::array([
          json_number(column),
          json_number(
            dimension.width.map(@xlsx2html.col_width_px).unwrap_or(96),
          ),
          json_number(id),
        ]),
      )
    }
    let panes = sheet.get_panes()
    sheets[name] = Json::object({
      "rows": json_number(rows),
      "columns": json_number(columns),
      "cells": Json::array(cells),
      "merges": Json::array(merges),
      "row_dimensions": Json::array(row_dimensions),
      "column_dimensions": Json::array(column_dimensions),
      "freeze_rows": json_number(if panes.freeze { panes.y_split } else { 0 }),
      "freeze_columns": json_number(
        if panes.freeze {
          panes.x_split
        } else {
          0
        },
      ),
    })
  }
  Json::object({
    "sheets": Json::object(sheets),
    "styles": Json::object(styles),
  })
}

///|
fn needs_virtual_excel(
  workbook : @xlsx.Workbook,
  rows : Int32,
  columns : Int32,
) -> Bool {
  let mut total = 0
  for sheet in workbook.sheets() {
    let (height, width) = virtual_sheet_bounds(sheet)
    if height > rows || width > columns || (width > 0 && height > 32768 / width) {
      return true
    }
    total += height * width
    if total > 32768 {
      return true
    }
  }
  false
}

///|
fn virtual_excel_shell(rendered : String, model : Json) -> String {
  let mut html = extract_html_body(rendered)
  if model is Object(fields) && fields.get("sheets") is Some(Object(sheets)) {
    for name, _ in sheets {
      let marker = "<section data-sheet=\"" +
        @xlsx2html.escape_html_attribute(name) +
        "\">"
      guard html.find(marker) is Some(start) else { continue }
      let begin = start + marker.length()
      guard html[begin:].find("</section>") is Some(length) else { continue }
      let section = html[begin:begin + length].to_owned()
      let section = match (section.find("<table>"), section.find("</table>")) {
        (Some(a), Some(b)) =>
          section[:a].to_owned() + section[b + 8:].to_owned()
        _ => section
      }
      html = html[:begin].to_owned() +
        section.replace(old="<p>(empty sheet)</p>", new="") +
        "<table class=\"office-virtual-table\" data-office-grid-decorated=\"true\"><tbody></tbody></table>" +
        html[begin + length:].to_owned()
    }
  }
  while html.find("<p class=\"truncated\">") is Some(start) {
    guard html[start:].find("</p>") is Some(end) else { break }
    html = html[:start].to_owned() + html[start + end + 4:].to_owned()
  }
  html +
  "<script type=\"application/json\" class=\"office-virtual-workbook\">" +
  model.stringify().split("<").join("\\u003c") +
  "</script>"
}

///|
fn virtual_sheet_bounds(sheet : @xlsx.Worksheet) -> (Int32, Int32) {
  let mut rows = sheet.max_row()
  let mut columns = sheet.max_col()
  let anchors : Array[String] = []
  for range in sheet.merged_cells() {
    if range.split(":").last() is Some(end) {
      anchors.push(end.to_owned())
    }
  }
  for note in sheet.comments() {
    anchors.push(note.cell)
  }
  for image in sheet.images() {
    anchors.push(image.reference)
  }
  for chart in sheet.charts() {
    anchors.push(chart.reference)
  }
  for anchor in anchors {
    let (column, row) = @xlsx.cell_name_to_coordinates(anchor) catch {
      _ => (0, 0)
    }
    rows = Int32::max(rows, row)
    columns = Int32::max(columns, column)
  }
  for (row, _) in sheet.materialized_row_dimensions() {
    rows = Int32::max(rows, row)
  }
  for (column, _) in sheet.materialized_col_dimensions() {
    columns = Int32::max(columns, column)
  }
  (rows, columns)
}
