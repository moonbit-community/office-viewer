///|
priv struct CellBounds {
  row : Int32
  col : Int32
  row_end : Int32
  col_end : Int32
}

///|
priv struct CellSelection {
  cells : Array[@dom.Element]
  top : Int32
  bottom : Int32
  left : Int32
  right : Int32
}

///|
fn cell_bounds(cell : @dom.Element) -> CellBounds? {
  let row = numeric_attr(cell, "data-row")
  let col = numeric_attr(cell, "data-col")
  if !cell.has_attribute("data-row") ||
    !cell.has_attribute("data-col") ||
    !finite(row) ||
    !finite(col) {
    return None
  }
  let row = row.to_int()
  let col = col.to_int()
  Some({
    row,
    col,
    row_end: row + Int32::max(1, numeric_attr(cell, "rowspan").to_int()) - 1,
    col_end: col + Int32::max(1, numeric_attr(cell, "colspan").to_int()) - 1,
  })
}

///|
fn clear_cells(table : @dom.Element) -> Unit {
  for cell in all(table, "td.office-cell-selected, td.office-cell-active") {
    toggle_class(cell, "office-cell-selected", false)
    toggle_class(cell, "office-cell-active", false)
  }
}

///|
fn selected_range(selected : CellSelection) -> (String, String, String) {
  let start = column_name(selected.left) + selected.top.to_string()
  let end = column_name(selected.right) + selected.bottom.to_string()
  (start, end, if start == end { start } else { start + ":" + end })
}

///|
fn SurfaceRuntime::select_cells(
  self : SurfaceRuntime,
  anchor : @dom.Element,
  active : @dom.Element,
) -> CellSelection? {
  guard anchor.closest("table") is Some(table) &&
    active.closest("table") is Some(active_table) &&
    same(table, active_table) else {
    return None
  }
  guard cell_bounds(anchor) is Some(a) && cell_bounds(active) is Some(b) else {
    return None
  }
  let top = Int32::min(a.row, b.row)
  let bottom = Int32::max(a.row_end, b.row_end)
  let left = Int32::min(a.col, b.col)
  let right = Int32::max(a.col_end, b.col_end)
  let cells : Array[@dom.Element] = []
  clear_cells(table)
  for cell in all(table, "td[data-cell]") {
    if cell_bounds(cell) is Some(coordinates) {
      if coordinates.row <= bottom &&
        coordinates.row_end >= top &&
        coordinates.col <= right &&
        coordinates.col_end >= left {
        toggle_class(cell, "office-cell-selected", true)
        cells.push(cell)
      }
    }
  }
  toggle_class(active, "office-cell-active", true)
  let selected : CellSelection = { cells, top, bottom, left, right, }
  self.update_formula_bar(Some(active), selected_range(selected).2)
  cells.stable_sort_by((left, right) => {
    let a = cell_bounds(left).unwrap()
    let b = cell_bounds(right).unwrap()
    if a.row == b.row {
      a.col - b.col
    } else {
      a.row - b.row
    }
  })
  Some(selected)
}

///|
fn SurfaceRuntime::excel_selection(
  self : SurfaceRuntime,
  pointer : CellPointer,
) -> Unit {
  guard self.select_cells(pointer.anchor, pointer.active) is Some(selected) &&
    !selected.cells.is_empty() else {
    return
  }
  let sheet = pointer.table
    .closest("section")
    .map(section => attr(section, "data-sheet"))
    .unwrap_or(attr(selected.cells[0], "data-sheet"))
  let (start, end, range) = selected_range(selected)
  let cells = selected.cells.map(cell => {
    let coordinates = cell_bounds(cell).unwrap()
    Json::object({
      "address": Json::string(attr(cell, "data-cell")),
      "row": number(coordinates.row.to_double()),
      "column": number(coordinates.col.to_double()),
      "text": Json::string(text(cell)),
    })
  })
  let leaf = if start == end {
    "cell[" + start + "]"
  } else {
    "range[" + range + "]"
  }
  self.emit(
    Json::object({
      "schema": Json::string("office.selection.v1"),
      "file": Json::string(self.document_file),
      "format": Json::string("xlsx"),
      "kind": Json::string("excel"),
      "sheet": Json::string(sheet),
      "range": Json::string(range),
      "selector": Json::string(
        "/xlsx/sheet[name=" + Json::string(sheet).stringify() + "]/" + leaf,
      ),
      "stability": Json::string("snapshot-relative"),
      "cells": Json::array(cells),
    }),
  )
}
