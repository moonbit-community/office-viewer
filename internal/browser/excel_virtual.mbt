///|
priv struct SheetAxis {
  count : Int32
  size : Double
  changes : Array[(Int32, Double)]
  styles : Map[Int32, Int32]
}

///|
fn sheet_axis(count : Int32, size : Double, dimensions : Json) -> SheetAxis {
  let changes : Array[(Int32, Double)] = []
  let styles : Map[Int32, Int32] = Map([])
  for entry in array(dimensions) {
    if array(entry) is [Number(index, ..), Number(value, ..), Number(id, ..)] {
      let index = index.to_int()
      if index > 0 && index <= count {
        changes.push((index, Double::max(1.0, value) - size))
        if id >= 0.0 {
          styles[index] = id.to_int()
        }
      }
    }
  }
  changes.sort_by((a, b) => a.0 - b.0)
  let mut sum = 0.0
  for i, change in changes {
    sum += change.1
    changes[i] = (change.0, sum)
  }
  { count, size, changes, styles, }
}

///|
fn SheetAxis::offset(self : SheetAxis, count : Int32) -> Double {
  let mut low = 0
  let mut high = self.changes.length()
  while low < high {
    let mid = low + (high - low) / 2
    if self.changes[mid].0 <= count {
      low = mid + 1
    } else {
      high = mid
    }
  }
  count.to_double() * self.size +
  (if low > 0 { self.changes[low - 1].1 } else { 0.0 })
}

///|
fn SheetAxis::index_at(self : SheetAxis, offset : Double) -> Int32 {
  let mut low = 1
  let mut high = self.count
  while low < high {
    let mid = low + (high - low) / 2
    if self.offset(mid) <= offset {
      low = mid + 1
    } else {
      high = mid
    }
  }
  low
}

///|
priv struct LazySheet {
  name : String
  section : @dom.Element
  viewport : @dom.Element
  table : @dom.Element
  body : @dom.Element
  rows : SheetAxis
  columns : SheetAxis
  heading_width : Double
  cells : Map[(Int32, Int32), Json]
  records : Array[Json]
  merges : Array[CellBounds]
  merge_ends : Array[Int32]
  frozen_rows : Int32
  frozen_columns : Int32
  mut selection : CellBounds?
  mut frame : Double?
}

///|
fn lazy_integer(value : Json) -> Int32 {
  match value {
    Number(n, ..) => n.to_int()
    _ => 0
  }
}

///|
fn lazy_cell(value : Json, index : Int32) -> Json {
  array(value).get(index).unwrap_or(Json::null())
}

///|
fn lazy_extent(axis : SheetAxis, heading : Double) -> Double {
  Double::min(8000000.0, axis.offset(axis.count) + heading)
}

///|
fn lazy_scroll(
  axis : SheetAxis,
  heading : Double,
  view : Double,
  scroll : Double,
) -> Double {
  let maximum = Double::max(0.0, axis.offset(axis.count) + heading - view)
  let physical = Double::max(1.0, lazy_extent(axis, heading) - view)
  Double::min(maximum, Double::max(0.0, scroll / physical * maximum))
}

///|
fn lazy_position(
  axis : SheetAxis,
  heading : Double,
  view : Double,
  logical : Double,
) -> Double {
  let maximum = Double::max(1.0, axis.offset(axis.count) + heading - view)
  Double::max(
    0.0,
    Double::min(
      lazy_extent(axis, heading) - view,
      logical / maximum * (lazy_extent(axis, heading) - view),
    ),
  )
}

///|
fn SurfaceRuntime::prepare_lazy_excel(self : SurfaceRuntime) -> Unit {
  let model = metadata(self.host, ".office-virtual-workbook")
  self.lazy_styles = field(model, "styles")
  for section in all(self.host, "section[data-sheet]") {
    guard section.query_selector("table.office-virtual-table") is Some(table) else {
      continue
    }
    let name = attr(section, "data-sheet")
    let data = field(field(model, "sheets"), name)
    guard data is Object(_) else { continue }
    let records = array(field(data, "cells"))
    let cells : Map[(Int32, Int32), Json] = Map([])
    for cell in records {
      cells[(lazy_integer(lazy_cell(cell, 0)), lazy_integer(lazy_cell(cell, 1)))] = cell
    }
    let rows = sheet_axis(
      lazy_integer(field(data, "rows")),
      26.0,
      field(data, "row_dimensions"),
    )
    let columns = sheet_axis(
      lazy_integer(field(data, "columns")),
      96.0,
      field(data, "column_dimensions"),
    )
    let merges : Array[CellBounds] = []
    for entry in array(field(data, "merges")) {
      if array(entry) is [a, b, c, d] {
        merges.push({
          row: lazy_integer(a),
          col: lazy_integer(b),
          row_end: lazy_integer(c),
          col_end: lazy_integer(d),
        })
      }
    }
    merges.sort_by((a, b) => {
      if a.row == b.row {
        a.col - b.col
      } else {
        a.row - b.row
      }
    })
    let merge_ends : Array[Int32] = []
    let mut end = 0
    for area in merges {
      end = Int32::max(end, area.row_end)
      merge_ends.push(end)
    }
    let heading_width = Double::max(
      44.0,
      rows.count.to_string().length().to_double() * 8.0 + 14.0,
    )
    let viewport = element("div", class_name="office-sheet-viewport")
    style(viewport, "overflow", "auto")
    style(viewport, "max-height", "65vh")
    style(viewport, "max-width", "100%")
    style(viewport, "position", "relative")
    let grid = element("div", class_name="office-sheet-grid")
    style(grid, "position", "relative")
    style(grid, "width", "max-content")
    if table.get_parent_element() is Some(parent) {
      insert_before(parent, viewport, table)
    }
    append(viewport, grid)
    append(grid, table)
    if section.query_selector(":scope > .sheet-objects") is Some(objects) {
      append(grid, objects)
    }
    for
      (key, value) in [
        ("display", "block"),
        ("position", "relative"),
        ("border-collapse", "separate"),
        ("margin", "0"),
        ("overflow", "hidden"),
      ] {
      style(table, key, value)
    }
    px(table, "width", lazy_extent(columns, heading_width))
    px(table, "height", lazy_extent(rows, 26.0))
    table.set_attribute("role", "grid")
    table.set_attribute("aria-rowcount", rows.count.to_string())
    table.set_attribute("aria-colcount", columns.count.to_string())
    let body = table.query_selector("tbody").unwrap()
    style(body, "display", "block")
    let sheet : LazySheet = {
      name,
      section,
      viewport,
      table,
      body,
      rows,
      columns,
      heading_width,
      cells,
      records,
      merges,
      merge_ends,
      frozen_rows: lazy_integer(field(data, "freeze_rows")),
      frozen_columns: lazy_integer(field(data, "freeze_columns")),
      selection: None,
      frame: None,
    }
    self.lazy_sheets[name] = sheet
    self.listen(viewport.as_event_target(), "scroll", _ => {
      self.schedule_lazy_sheet(sheet)
    })
    self.cleanups.push(
      observe_resize(viewport, () => self.schedule_lazy_sheet(sheet)),
    )
  }
  if self.host.query_selector(".office-virtual-workbook") is Some(script) {
    remove(script)
  }
}

///|
fn SurfaceRuntime::schedule_lazy_sheet(
  self : SurfaceRuntime,
  sheet : LazySheet,
) -> Unit {
  if !self.alive || sheet.frame is Some(_) {
    return
  }
  sheet.frame = Some(
    @dom.window().request_animation_frame(_ => {
      sheet.frame = None
      if self.alive {
        self.draw_lazy_sheet(sheet)
      }
    }),
  )
}

///|
fn lazy_indices(
  axis : SheetAxis,
  start : Double,
  size : Double,
  frozen : Int32,
) -> Array[Int32] {
  let indices : Array[Int32] = []
  if axis.count == 0 {
    return indices
  }
  let frozen = Int32::min(frozen, axis.index_at(size))
  for index in 1..<=frozen {
    indices.push(index)
  }
  let first = Int32::max(frozen + 1, axis.index_at(Double::max(0.0, start)) - 2)
  let last = Int32::min(axis.count, axis.index_at(start + size) + 2)
  for index in first..<=last {
    indices.push(index)
  }
  indices
}

///|
fn lazy_merge(sheet : LazySheet, row : Int32, col : Int32) -> CellBounds? {
  let mut low = 0
  let mut high = sheet.merges.length()
  while low < high {
    let mid = low + (high - low) / 2
    if sheet.merges[mid].row <= row {
      low = mid + 1
    } else {
      high = mid
    }
  }
  while low > 0 {
    low -= 1
    if sheet.merge_ends[low] < row {
      break
    }
    let area = sheet.merges[low]
    if area.row_end >= row && area.col <= col && area.col_end >= col {
      return Some(area)
    }
  }
  None
}

///|
fn overlaps(a : CellBounds, b : CellBounds) -> Bool {
  a.row <= b.row_end &&
  a.row_end >= b.row &&
  a.col <= b.col_end &&
  a.col_end >= b.col
}

///|
fn SurfaceRuntime::draw_lazy_sheet(
  self : SurfaceRuntime,
  sheet : LazySheet,
) -> Unit {
  if get_boolean(sheet.section, "hidden") {
    return
  }
  let view_height = sheet.viewport.get_client_height()
  let view_width = sheet.viewport.get_client_width()
  if view_height <= 0.0 || view_width <= 0.0 {
    return
  }
  let sy = sheet.viewport.get_scroll_top()
  let sx = sheet.viewport.get_scroll_left()
  let y = lazy_scroll(sheet.rows, 26.0, view_height, sy)
  let x = lazy_scroll(sheet.columns, sheet.heading_width, view_width, sx)
  let rows = lazy_indices(sheet.rows, y, view_height, sheet.frozen_rows)
  let columns = lazy_indices(sheet.columns, x, view_width, sheet.frozen_columns)
  sheet.table.set_attribute(
    "data-window-start-row",
    rows.get(0).unwrap_or(0).to_string(),
  )
  sheet.body.set_inner_html("")
  let seen : Map[(Int32, Int32), Bool] = Map([])
  let rendered_cells : Map[String, @dom.Element] = Map([])
  for row in rows {
    let tr = element("tr")
    style(tr, "display", "contents")
    append(sheet.body, tr)
    let heading = element(
      "th",
      class_name="office-row-heading",
      text=row.to_string(),
    )
    place_lazy_cell(
      sheet,
      heading,
      { row, col: 0, row_end: row, col_end: 0, },
      x,
      y,
      sx,
      sy,
    )
    append(tr, heading)
    for col in columns {
      let bounds = lazy_merge(sheet, row, col).unwrap_or({
        row,
        col,
        row_end: row,
        col_end: col,
      })
      if seen.contains((bounds.row, bounds.col)) {
        continue
      }
      seen[(bounds.row, bounds.col)] = true
      let data = sheet.cells
        .get((bounds.row, bounds.col))
        .unwrap_or(Json::null())
      let address = column_name(bounds.col) + bounds.row.to_string()
      let cell = element("td", text=string(lazy_cell(data, 2)))
      cell.set_attribute("data-office-cell-value", string(lazy_cell(data, 2)))
      let id = if data is Null {
        sheet.rows.styles
        .get(bounds.row)
        .otherwise(() => sheet.columns.styles.get(bounds.col))
        .unwrap_or(0)
      } else {
        lazy_integer(lazy_cell(data, 3))
      }
      cell.set_attribute(
        "style",
        string(field(self.lazy_styles, id.to_string())),
      )
      for
        (key, value) in [
          ("data-sheet", sheet.name),
          ("data-row", bounds.row.to_string()),
          ("data-col", bounds.col.to_string()),
          ("data-cell", address),
          ("role", "gridcell"),
          ("aria-rowindex", bounds.row.to_string()),
          ("aria-colindex", bounds.col.to_string()),
        ] {
        cell.set_attribute(key, value)
      }
      if bounds.row_end > bounds.row {
        cell.set_attribute(
          "rowspan",
          (bounds.row_end - bounds.row + 1).to_string(),
        )
      }
      if bounds.col_end > bounds.col {
        cell.set_attribute(
          "colspan",
          (bounds.col_end - bounds.col + 1).to_string(),
        )
      }
      let formula = string(lazy_cell(data, 4))
      if formula != "" {
        cell.set_attribute("data-formula", formula)
      }
      toggle_class(
        cell,
        "formula-pending",
        lazy_cell(data, 5) == Json::boolean(true),
      )
      if sheet.selection is Some(selected) {
        toggle_class(cell, "office-cell-selected", overlaps(bounds, selected))
      }
      if self.lazy_query != "" {
        toggle_class(
          cell,
          "office-search-hit",
          locale_lowercase(text(cell)).contains(self.lazy_query),
        )
      }
      place_lazy_cell(sheet, cell, bounds, x, y, sx, sy)
      append(tr, cell)
      rendered_cells[address] = cell
    }
  }
  let head = element("tr", class_name="office-grid-column-headings")
  style(head, "display", "contents")
  append(sheet.body, head)
  for col in columns {
    let heading = element(
      "th",
      class_name="office-column-heading",
      text=column_name(col),
    )
    place_lazy_cell(
      sheet,
      heading,
      { row: 0, col, row_end: 0, col_end: col, },
      x,
      y,
      sx,
      sy,
    )
    append(head, heading)
  }
  let corner = element("th", class_name="office-grid-corner")
  place_lazy_cell(
    sheet,
    corner,
    { row: 0, col: 0, row_end: 0, col_end: 0, },
    x,
    y,
    sx,
    sy,
  )
  append(head, corner)
  for item in self.annotations {
    if string(field(item.anchor, "sheet")) == sheet.name {
      item.targets.clear()
      if rendered_cells.get(string(field(item.anchor, "cell"))) is Some(target) {
        item.targets.push(target)
        toggle_class(target, "office-has-annotation", true)
        let marker = element(
          "button",
          class_name="office-note-marker",
          text="1",
        )
        marker.set_attribute("type", "button")
        marker.set_attribute("aria-label", "Note by " + author_name(item.data))
        marker.add_event_listener("click", event => {
          event.prevent_default()
          event.stop_propagation()
          if self.show_annotation(item) {
            ignore(self.select_annotation(item.data))
          }
        })
        append(target, marker)
      }
    }
  }
  self.layout_lazy_objects(sheet)
  self.schedule_layout()
}

///|
fn place_lazy_cell(
  sheet : LazySheet,
  cell : @dom.Element,
  bounds : CellBounds,
  x : Double,
  y : Double,
  sx : Double,
  sy : Double,
) -> Unit {
  let left = if bounds.col == 0 {
    0.0
  } else {
    sheet.heading_width + sheet.columns.offset(bounds.col - 1)
  }
  let top = if bounds.row == 0 {
    0.0
  } else {
    26.0 + sheet.rows.offset(bounds.row - 1)
  }
  let pinned_row = bounds.row <= sheet.frozen_rows
  let pinned_col = bounds.col <= sheet.frozen_columns
  let left = sx + (if pinned_col { left } else { left - x })
  let top = sy + (if pinned_row { top } else { top - y })
  px(cell, "left", left)
  px(cell, "top", top)
  px(
    cell,
    "width",
    if bounds.col == 0 {
      sheet.heading_width
    } else {
      sheet.columns.offset(bounds.col_end) -
      sheet.columns.offset(bounds.col - 1)
    },
  )
  px(
    cell,
    "height",
    if bounds.row == 0 {
      26.0
    } else {
      sheet.rows.offset(bounds.row_end) - sheet.rows.offset(bounds.row - 1)
    },
  )
  for
    (key, value) in [
      ("position", "absolute"),
      ("display", "block"),
      ("box-sizing", "border-box"),
      ("min-width", "0"),
      ("min-height", "0"),
      ("overflow", "hidden"),
      ("padding", "2px 6px"),
    ] {
    style(cell, key, value)
  }
  style(
    cell,
    "z-index",
    if pinned_row && pinned_col {
      "7"
    } else if pinned_row || pinned_col {
      "5"
    } else {
      "1"
    },
  )
  if computed_style(cell, "background-color") == "rgba(0, 0, 0, 0)" {
    style(cell, "background-color", "#fff")
  }
}

///|
fn SurfaceRuntime::reveal_lazy_cell(
  self : SurfaceRuntime,
  name : String,
  row : Int32,
  column : Int32,
) -> Bool {
  guard self.lazy_sheets.get(name) is Some(sheet) &&
    row >= 1 &&
    row <= sheet.rows.count &&
    column >= 1 &&
    column <= sheet.columns.count else {
    return false
  }
  if attr(self.host, "data-office-active-sheet") != name {
    match
      all(self.host, "[data-excel-sheet-tab]")
      .iter()
      .find_first(tab => attr(tab, "data-excel-sheet-tab") == name) {
      Some(tab) => if tab.to_html_element() is Some(element) { element.click() }
      None => self.set_excel_sheet(name)
    }
  }
  sheet.viewport.set_scroll_top(
    lazy_position(
      sheet.rows,
      26.0,
      sheet.viewport.get_client_height(),
      Double::max(
        0.0,
        sheet.rows.offset(row - 1) -
        sheet.rows.offset(Int32::min(sheet.frozen_rows, row - 1)),
      ),
    ),
  )
  sheet.viewport.set_scroll_left(
    lazy_position(
      sheet.columns,
      sheet.heading_width,
      sheet.viewport.get_client_width(),
      Double::max(
        0.0,
        sheet.columns.offset(column - 1) -
        sheet.columns.offset(Int32::min(sheet.frozen_columns, column - 1)),
      ),
    ),
  )
  self.draw_lazy_sheet(sheet)
  true
}

///|
fn SurfaceRuntime::search_lazy_excel(
  self : SurfaceRuntime,
  query : String,
) -> Unit {
  let needle = locale_lowercase(query.trim().to_owned())
  if needle == self.lazy_query {
    return
  }
  self.lazy_query = needle
  if needle != "" {
    for sheet in self.lazy_sheets.values() {
      if sheet.records
        .iter()
        .find_first(cell => {
          let row = lazy_integer(lazy_cell(cell, 0))
          let col = lazy_integer(lazy_cell(cell, 1))
          let matched = locale_lowercase(string(lazy_cell(cell, 2))).contains(
            needle,
          )
          matched &&
          (match lazy_merge(sheet, row, col) {
            Some(area) => area.row == row && area.col == col
            None => true
          })
        })
        is Some(cell) {
        ignore(
          self.reveal_lazy_cell(
            sheet.name,
            lazy_integer(lazy_cell(cell, 0)),
            lazy_integer(lazy_cell(cell, 1)),
          ),
        )
        return
      }
    }
  }
  for sheet in self.lazy_sheets.values() {
    self.draw_lazy_sheet(sheet)
  }
}

///|
fn SurfaceRuntime::lazy_selection_cells(
  self : SurfaceRuntime,
  name : String,
  bounds : CellBounds,
) -> Array[Json]? {
  guard self.lazy_sheets.get(name) is Some(sheet) else { return None }
  let cells : Array[Json] = []
  for data in sheet.records {
    let row = lazy_integer(lazy_cell(data, 0))
    let col = lazy_integer(lazy_cell(data, 1))
    if row >= bounds.row &&
      row <= bounds.row_end &&
      col >= bounds.col &&
      col <= bounds.col_end {
      if lazy_merge(sheet, row, col) is Some(area) &&
        (area.row != row || area.col != col) {
        continue
      }
      cells.push(
        Json::object({
          "address": Json::string(column_name(col) + row.to_string()),
          "row": number(row.to_double()),
          "column": number(col.to_double()),
          "text": lazy_cell(data, 2),
        }),
      )
    }
  }
  Some(cells)
}

///|
fn lazy_address(address : String) -> (Int32, Int32)? {
  let mut column = 0
  let mut row = 0
  let mut digits = false
  for unit in address {
    if !digits && unit >= 'A' && unit <= 'Z' {
      column = column * 26 + unit.to_int() - 64
    } else if unit >= '0' && unit <= '9' {
      digits = true
      row = row * 10 + unit.to_int() - 48
    } else {
      return None
    }
    if row > 1048576 || column > 16384 {
      return None
    }
  }
  if row > 0 && column > 0 {
    Some((row, column))
  } else {
    None
  }
}

///|
fn SurfaceRuntime::layout_lazy_objects(
  self : SurfaceRuntime,
  sheet : LazySheet,
) -> Unit {
  guard !get_boolean(sheet.section, "hidden") &&
    sheet.section.query_selector(".sheet-objects") is Some(overlay) else {
    return
  }
  let objects = array(
    field(
      field(metadata(self.host, ".office-object-metadata"), "sheets"),
      sheet.name,
    ),
  )
  let sy = sheet.viewport.get_scroll_top()
  let sx = sheet.viewport.get_scroll_left()
  let h = sheet.viewport.get_client_height()
  let w = sheet.viewport.get_client_width()
  let y = lazy_scroll(sheet.rows, 26.0, h, sy)
  let x = lazy_scroll(sheet.columns, sheet.heading_width, w, sx)
  toggle_class(overlay, "office-object-overlay", true)
  for
    index, figure in all(
      overlay, "figure.image, figure.chart, figure.chart-placeholder",
    ) {
    guard objects.get(index) is Some(info) &&
      lazy_address(string(field(info, "anchor"))) is Some((row, col)) else {
      continue
    }
    let left = sx +
      sheet.heading_width +
      sheet.columns.offset(col - 1) -
      (if col <= sheet.frozen_columns { 0.0 } else { x }) +
      object_number(info, "offset_x", 0.0)
    let top = sy +
      26.0 +
      sheet.rows.offset(row - 1) -
      (if row <= sheet.frozen_rows { 0.0 } else { y }) +
      object_number(info, "offset_y", 0.0)
    let width = object_number(info, "width_px", 1.0)
    let height = object_number(info, "height_px", 1.0)
    set_boolean(
      figure,
      "hidden",
      left > sx + w || top > sy + h || left + width < sx || top + height < sy,
    )
    for
      (key, value) in [
        ("position", "absolute"),
        ("box-sizing", "border-box"),
        ("margin", "0"),
        ("pointer-events", "none"),
      ] {
      style(figure, key, value)
    }
    px(figure, "left", left)
    px(figure, "top", top)
    px(figure, "width", width)
    px(figure, "height", height)
    if figure.query_selector("img") is Some(image) {
      for
        (key, value) in [
          ("display", "block"),
          ("width", "100%"),
          ("height", "100%"),
          ("object-fit", "contain"),
        ] {
        style(image, key, value)
      }
    }
  }
}
