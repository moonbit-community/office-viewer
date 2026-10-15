///|
fn SurfaceRuntime::update_formula_bar(
  self : SurfaceRuntime,
  cell : @dom.Element?,
  range : String,
) -> Unit {
  if self.kind != "excel" {
    return
  }
  if self.host.query_selector("[data-excel-name]") is Some(node) {
    set_text(
      node,
      if range != "" {
        range
      } else {
        cell.map(cell => attr(cell, "data-cell")).unwrap_or("")
      },
    )
  }
  if self.host.query_selector("[data-excel-value]") is Some(node) {
    match cell {
      None => {
        set_text(node, "Select a cell to inspect its value")
        toggle_class(node, "is-placeholder", true)
      }
      Some(cell) => {
        let formula = attr(cell, "data-formula")
        let value = if formula != "" { "=" + formula } else { text(cell) }
        set_text(node, if value == "" { " " } else { value })
        toggle_class(node, "is-placeholder", false)
      }
    }
  }
}

///|
fn SurfaceRuntime::set_excel_sheet(
  self : SurfaceRuntime,
  sheet : String,
) -> Unit {
  if self.kind != "excel" {
    return
  }
  self.clear_context()
  self.cancel_pointer()
  for section in all(self.host, "section[data-sheet]") {
    let active = attr(section, "data-sheet") == sheet
    toggle_class(section, "is-active", active)
    set_boolean(section, "hidden", !active)
    if active {
      if section.query_selector("table") is Some(table) {
        table.set_scroll_left(0.0)
      }
    }
  }
  for tab in all(self.host, "[data-excel-sheet-tab]") {
    let active = attr(tab, "data-excel-sheet-tab") == sheet
    toggle_class(tab, "is-active", active)
    tab.set_attribute("aria-selected", active.to_string())
  }
  self.host.set_attribute("data-office-active-sheet", sheet)
  for cell in all(self.host, "td.office-cell-selected, td.office-cell-active") {
    toggle_class(cell, "office-cell-selected", false)
    toggle_class(cell, "office-cell-active", false)
  }
  self.update_formula_bar(None, "")
  self.layout_excel_objects()
  call_hook(self.host, "__officeFreezeLayout")
  if self.rail is Some(rail) {
    rail.set_scroll_top(0.0)
  }
  self.schedule_layout()
}

///|
fn SurfaceRuntime::prepare_excel(self : SurfaceRuntime) -> Unit {
  if self.kind != "excel" {
    return
  }
  let formulas = field(
    metadata(self.host, ".office-formula-metadata"),
    "sheets",
  )
  for section in all(self.host, "section[data-sheet]") {
    let sheet = attr(section, "data-sheet")
    for table in all(section, "table") {
      let occupied : Map[(Int32, Int32), Bool] = Map([])
      for row_index, row_element in all(table, "tr") {
        let mut col = 1
        for cell in all(row_element, "td") {
          while occupied.contains((row_index, col)) {
            col += 1
          }
          let row = row_index + 1
          let colspan = Int32::max(1, numeric_attr(cell, "colspan").to_int())
          let rowspan = Int32::max(1, numeric_attr(cell, "rowspan").to_int())
          let address = column_name(col) + row.to_string()
          for
            (key, value) in [
              ("data-sheet", sheet),
              ("data-row", row.to_string()),
              ("data-col", col.to_string()),
              ("data-cell", address),
            ] {
            cell.set_attribute(key, value)
          }
          if field(field(formulas, sheet), address) is String(formula) {
            cell.set_attribute("data-formula", formula)
          }
          for r in row_index..<(row_index + rowspan) {
            for c in col..<(col + colspan) {
              occupied[(r, c)] = true
            }
          }
          col += colspan
        }
      }
      decorate_excel_grid(table)
    }
  }
  if self.host.query_selector(".office-formula-metadata") is Some(node) {
    remove(node)
  }
  self.enhance_excel_surface()
  self.layout_excel_objects()
}

///|
fn decorate_excel_grid(table : @dom.Element) -> Unit {
  if attr(table, "data-office-grid-decorated") == "true" {
    return
  }
  let colgroup = table.query_selector("colgroup")
  let mut column_count = colgroup
    .map(group => all(group, "col").length())
    .unwrap_or(0)
  if column_count == 0 {
    for row in all(table, "tr") {
      let mut columns = 0
      for cell in all(row, "td") {
        columns += Int32::max(1, numeric_attr(cell, "colspan").to_int())
      }
      column_count = Int32::max(column_count, columns)
    }
  }
  if colgroup is Some(group) {
    let col = element("col", class_name="office-row-header-column")
    style(col, "width", "44px")
    group.insert_before(
      col.as_node(),
      @js.Nullable::from_option(
        group.query_selector("col").map(node => node.as_node()),
      ),
    )
  }
  let head = element("thead", class_name="office-grid-column-headings")
  let head_row = element("tr")
  let corner = element("th", class_name="office-grid-corner")
  corner.set_attribute("aria-label", "Select all cells")
  append(head_row, corner)
  for column in 1..<=column_count {
    let heading = element(
      "th",
      class_name="office-column-heading",
      text=column_name(column),
    )
    heading.set_attribute("scope", "col")
    append(head_row, heading)
  }
  append(head, head_row)
  if table.query_selector("tbody") is Some(body) {
    for index, row in all(body, "tr") {
      let row_number = row
        .query_selector("td[data-row]")
        .map(cell => attr(cell, "data-row"))
        .unwrap_or((index + 1).to_string())
      let heading = element(
        "th",
        class_name="office-row-heading",
        text=row_number,
      )
      heading.set_attribute("scope", "row")
      row.insert_before(
        heading.as_node(),
        @js.Nullable::from_option(
          row.query_selector("td").map(node => node.as_node()),
        ),
      )
    }
    insert_before(table, head, body)
  } else {
    append(table, head)
  }
  for cell in all(table, "td[data-cell]") {
    cell.set_attribute("role", "gridcell")
  }
  table.set_attribute("data-office-grid-decorated", "true")
}

///|
fn SurfaceRuntime::enhance_excel_surface(self : SurfaceRuntime) -> Unit {
  guard self.host.query_selector(".sheet-document") is Some(book) else {
    return
  }
  let sections = all(book, "section[data-sheet]")
  if sections.is_empty() {
    return
  }
  let workspace = element("div", class_name="excel-workspace")
  let toolbar = element("div", class_name="excel-view-toolbar")
  for
    (class_name, label) in [
      ("excel-view-title", "Spreadsheet"),
      ("excel-readonly-badge", "Read only"),
      ("excel-view-hint", "Select cells to give an Agent precise context"),
    ] {
    append(toolbar, element("span", class_name~, text=label))
  }
  let formula_bar = element("div", class_name="excel-formula-bar")
  let name_box = element("span", class_name="excel-name-box")
  name_box.set_attribute("data-excel-name", "")
  name_box.set_attribute("aria-label", "Cell reference")
  append(formula_bar, name_box)
  let fx = element("span", class_name="excel-fx", text="fx")
  fx.set_attribute("aria-hidden", "true")
  append(formula_bar, fx)
  let value_box = element(
    "span",
    class_name="excel-value-box is-placeholder",
    text="Select a cell to inspect its value",
  )
  value_box.set_attribute("data-excel-value", "")
  append(formula_bar, value_box)
  let tabs = element("nav", class_name="excel-sheet-tabs")
  tabs.set_attribute("aria-label", "Worksheets")
  append(tabs, element("span", class_name="excel-tabs-label", text="Sheets"))
  let tab_list = element("div", class_name="excel-tab-list")
  append(tabs, tab_list)
  toggle_class(book, "office-excel-workbook", true)
  if book.query_selector(":scope > h1") is Some(title) {
    toggle_class(title, "excel-document-title", true)
  }
  for section in sections {
    toggle_class(section, "excel-sheet", true)
    if section.query_selector(":scope > h2") is Some(heading) {
      toggle_class(heading, "excel-sheet-heading", true)
    }
    let name = attr(section, "data-sheet")
    let tab = element("button", class_name="excel-sheet-tab", text=name)
    tab.set_attribute("type", "button")
    tab.set_attribute("data-excel-sheet-tab", name)
    tab.set_attribute("role", "tab")
    tab.set_attribute("aria-selected", "false")
    append(tab_list, tab)
    self.listen(tab.as_event_target(), "click", event => {
      event.prevent_default()
      self.set_excel_sheet(name)
    })
  }
  guard book.get_parent_element() is Some(parent) else { return }
  insert_before(parent, workspace, book)
  for child in [toolbar, formula_bar, book, tabs] {
    append(workspace, child)
  }
  let remembered = attr(self.host, "data-office-active-sheet")
  let sheet = if sections.any(section => {
      attr(section, "data-sheet") == remembered
    }) {
    remembered
  } else {
    attr(sections[0], "data-sheet")
  }
  self.set_excel_sheet(sheet)
}

///|
fn object_number(value : Json, key : String, fallback : Double) -> Double {
  let n = match field(value, key) {
    Number(n, ..) => n
    String(s) => browser_number(s)
    _ => 0.0
  }
  if !finite(n) || n == 0.0 {
    fallback
  } else {
    n
  }
}

///|
fn SurfaceRuntime::layout_excel_objects(self : SurfaceRuntime) -> Unit {
  if self.kind != "excel" {
    return
  }
  let sheets = field(metadata(self.host, ".office-object-metadata"), "sheets")
  if sheets is Null {
    return
  }
  let zoom = self.host
    .query_selector(".sheet-document")
    .map(node => browser_number(computed_style(node, "zoom")))
    .unwrap_or(1.0)
  let zoom = if finite(zoom) && zoom != 0.0 { zoom } else { 1.0 }
  for section in all(self.host, "section[data-sheet]") {
    let objects = array(field(sheets, attr(section, "data-sheet")))
    guard section.query_selector("table") is Some(table) else { continue }
    let grid = section.query_selector(".office-sheet-grid").unwrap_or(section)
    guard grid.query_selector(":scope > .sheet-objects") is Some(overlay) else {
      continue
    }
    toggle_class(section, "office-anchored-objects", true)
    toggle_class(overlay, "office-object-overlay", true)
    let origin = grid.get_bounding_client_rect()
    let cells = all(section, "td[data-cell]")
    let mut right = if width(table) != 0.0 {
      width(table)
    } else {
      table.get_bounding_client_rect().get_width()
    }
    let mut bottom = if height(table) != 0.0 {
      height(table)
    } else {
      table.get_bounding_client_rect().get_height()
    }
    for
      index, figure in all(
        overlay, "figure.image, figure.chart, figure.chart-placeholder",
      ) {
      guard objects.get(index) is Some(info) else { continue }
      let anchor = string(field(info, "anchor"))
      let anchor = if anchor == "" {
        attr(figure, "data-anchor")
      } else {
        anchor
      }
      guard cells.iter().find_first(cell => attr(cell, "data-cell") == anchor)
        is Some(cell) else {
        continue
      }
      let rect = cell.get_bounding_client_rect()
      let object_width = Double::max(1.0, object_number(info, "width_px", 1.0))
      let object_height = Double::max(
        1.0,
        object_number(info, "height_px", 1.0),
      )
      let left = (rect.get_left() - origin.get_left()) / zoom +
        object_number(info, "offset_x", 0.0)
      let top = (rect.get_top() - origin.get_top()) / zoom +
        object_number(info, "offset_y", 0.0)
      style(figure, "position", "absolute")
      px(figure, "left", left)
      px(figure, "top", top)
      px(figure, "width", object_width)
      px(figure, "height", object_height)
      style(figure, "box-sizing", "border-box")
      style(figure, "margin", "0")
      style(figure, "pointer-events", "none")
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
      if figure.query_selector("svg") is Some(svg) {
        style(svg, "display", "block")
        svg.set_attribute("width", object_width.to_string())
        svg.set_attribute("height", object_height.to_string())
        style(svg, "width", "100%")
        style(svg, "height", "100%")
      }
      right = Double::max(right, left + object_width)
      bottom = Double::max(bottom, top + object_height)
    }
    px(grid, "min-width", right.ceil())
    px(grid, "min-height", bottom.ceil())
  }
}
