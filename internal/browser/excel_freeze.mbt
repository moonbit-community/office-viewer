///|
priv struct FrozenSheet {
  section : @dom.Element
  table : @dom.Element
  rows : Double
  columns : Double
}

///|
fn freeze_count(
  settings : Json,
  section : @dom.Element,
  key : String,
  attribute : String,
) -> Double {
  let value = match settings {
    Null => numeric_attr(section, attribute)
    _ =>
      match field(settings, key) {
        Number(n, ..) => n
        String(s) => browser_number(s)
        Null => 0.0
        _ => @double.not_a_number
      }
  }
  if finite(value) &&
    value >= 0.0 &&
    value <= 9007199254740991.0 &&
    value.floor() == value {
    value
  } else {
    0.0
  }
}

///|
fn pin_cell(
  cell : @dom.Element,
  row : Int32,
  col : Int32,
  heading : Bool,
  sheet : FrozenSheet,
  left : Array[Double],
  top : Array[Double],
) -> Unit {
  let frozen_row = heading || row.to_double() <= sheet.rows
  let frozen_col = col == 0 || col.to_double() <= sheet.columns
  if !frozen_row && !frozen_col {
    return
  }
  style(cell, "position", "sticky")
  if frozen_row {
    px(cell, "top", if heading { 0.0 } else { top.get(row - 1).unwrap_or(0.0) })
  }
  if frozen_col {
    px(cell, "left", left.get(col).unwrap_or(0.0))
  }
  style(
    cell,
    "z-index",
    if frozen_row && frozen_col {
      "7"
    } else if frozen_row {
      "5"
    } else {
      "4"
    },
  )
  if computed_style(cell, "background-color") == "rgba(0, 0, 0, 0)" {
    style(cell, "background-color", "#fff")
  }
}

///|
fn layout_frozen_sheets(sheets : Array[FrozenSheet]) -> Unit {
  for sheet in sheets {
    if get_boolean(sheet.section, "hidden") || width(sheet.table) == 0.0 {
      continue
    }
    let headings = all(sheet.table, "thead th")
    let body_rows = all(sheet.table, "tbody tr")
    let left : Array[Double] = []
    let top : Array[Double] = []
    let mut x = 0.0
    let mut y = sheet.table.query_selector("thead").map(height).unwrap_or(0.0)
    for heading in headings {
      left.push(x)
      x += width(heading)
    }
    for row in body_rows {
      top.push(y)
      y += height(row)
    }
    for column, cell in headings {
      pin_cell(cell, 0, column, true, sheet, left, top)
    }
    for index, row in body_rows {
      if row.query_selector(".office-row-heading") is Some(heading) {
        pin_cell(heading, index + 1, 0, false, sheet, left, top)
      }
      for cell in all(row, "td[data-row][data-col]") {
        pin_cell(
          cell,
          numeric_attr(cell, "data-row").to_int(),
          numeric_attr(cell, "data-col").to_int(),
          false,
          sheet,
          left,
          top,
        )
      }
    }
  }
}

///|
pub fn install_excel_freeze(id : String) -> Unit {
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) &&
    has_class(host, "sheet-surface") else {
    return
  }
  call_hook(host, "__officeFreezeCleanup")
  let cleanups : Array[() -> Unit] = []
  let sheets : Array[FrozenSheet] = []
  let settings = field(metadata(host, ".office-freeze-metadata"), "sheets")
  if host.query_selector(".sheet-document") is Some(book) {
    style(book, "min-width", "0")
  }
  for section in all(host, "section[data-sheet]") {
    guard section.query_selector("table") is Some(table) else { continue }
    let viewport = section
      .query_selector(".office-sheet-viewport")
      .unwrap_or_else(() => {
        let viewport = element("div", class_name="office-sheet-viewport")
        for
          (key, value) in [
            ("overflow", "auto"),
            ("max-height", "65vh"),
            ("max-width", "100%"),
            ("position", "relative"),
          ] {
          style(viewport, key, value)
        }
        let grid = element("div", class_name="office-sheet-grid")
        for
          (key, value) in [
            ("position", "relative"),
            ("width", "max-content"),
            ("min-width", "100%"),
          ] {
          style(grid, key, value)
        }
        if table.get_parent_element() is Some(parent) {
          insert_before(parent, viewport, table)
        }
        append(viewport, grid)
        append(grid, table)
        if section.query_selector(":scope > .sheet-objects") is Some(objects) {
          append(grid, objects)
        }
        viewport
      })
    ignore(element_style(section).remove_property("min-width"))
    ignore(element_style(section).remove_property("min-height"))
    let settings = field(settings, attr(section, "data-sheet"))
    sheets.push({
      section,
      table,
      rows: freeze_count(settings, section, "rows", "data-freeze-rows"),
      columns: freeze_count(settings, section, "columns", "data-freeze-columns"),
    })
    let scroll = (_ : @dom.Event) => {
      call_hook(host, "__officeLayoutAnnotations")
    }
    viewport.add_event_listener("scroll", scroll)
    cleanups.push(() => viewport.remove_event_listener("scroll", scroll))
  }
  let layout = () => layout_frozen_sheets(sheets)
  set_callback(host, "__officeFreezeLayout", layout)
  cleanups.push(observe_resize(host, layout))
  let interaction_cleanup = get_callback(host, "__officeViewerCleanup").to_option()
  set_callback(host, "__officeFreezeCleanup", () => {
    for cleanup in cleanups {
      cleanup()
    }
    cleanups.clear()
    host.set_property(
      "__officeFreezeLayout",
      @js.Value::cast_from((@js.Nullable::null() : @js.Nullable[Unit])),
    )
    host.set_property(
      "__officeFreezeCleanup",
      @js.Value::cast_from((@js.Nullable::null() : @js.Nullable[Unit])),
    )
  })
  set_callback(host, "__officeViewerCleanup", () => {
    call_hook(host, "__officeFreezeCleanup")
    if interaction_cleanup is Some(cleanup) {
      cleanup()
    }
  })
  layout()
  call_hook(host, "__officeLayoutObjects")
}

///|
pub fn refresh_excel_freeze(id : String) -> Unit {
  if @dom.document().get_element_by_id(id).to_option() is Some(host) {
    call_hook(host, "__officeFreezeLayout")
  }
}
