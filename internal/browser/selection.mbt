///|
fn SurfaceRuntime::annotation_selection(self : SurfaceRuntime) -> Bool {
  if !self.alive {
    return false
  }
  guard native_selection().to_option() is Some(selection) else { return false }
  if selection_count(selection) != 1 || selection_collapsed(selection) {
    return false
  }
  let range = selection_range(selection)
  let start_node = range_node(range, true)
  let end_node = range_node(range, false)
  let start_body = node_element(start_node).bind(node => {
    node.closest("[data-office-annotation-body]")
  })
  let end_body = node_element(end_node).bind(node => {
    node.closest("[data-office-annotation-body]")
  })
  let intersects = self.cards.any(record => {
    range_intersects(range, record.card.as_node())
  })
  guard start_body is Some(body) &&
    end_body is Some(end_body) &&
    same(body, end_body) &&
    body.get_is_connected() else {
    return start_body is Some(_) || end_body is Some(_) || intersects
  }
  let known = if self.host.contains(body.as_node()) {
    self.bodies
    .iter()
    .find_first(pair => same(pair.0, body))
    .map(pair => pair.1)
  } else {
    None
  }
  let item = known.otherwise(() => {
    if attr(body, "data-office-annotation-surface") != self.id {
      return None
    }
    self.current_annotation(
      Json::object({
        "id": Json::string(attr(body, "data-office-annotation-body")),
        "document": Json::object({
          "file": Json::string(attr(body, "data-office-annotation-file")),
          "sha256": Json::string(attr(body, "data-office-annotation-sha256")),
        }),
      }),
    )
  })
  guard item is Some(item) && item.body is Some(contents) else { return true }
  let start = range_prefix_length(
    body.as_node(),
    start_node,
    range_offset(range, true),
    0,
  )
  let end = range_prefix_length(
    body.as_node(),
    end_node,
    range_offset(range, false),
    0,
  )
  if text(body) == contents && end > start {
    ignore(self.emit_annotation(item, start, end, "text"))
  }
  true
}

///|
priv struct WordAnchor {
  fields : Map[String, Json]
  page : Double
  top : Double
  left : Double
}

///|
fn SurfaceRuntime::word_selection(self : SurfaceRuntime) -> Unit {
  if !self.alive || self.annotation_selection() {
    return
  }
  guard native_selection().to_option() is Some(selection) else { return }
  if selection_count(selection) != 1 {
    return
  }
  let selected_text = selection_text(selection)
  if selected_text.trim().is_empty() {
    return
  }
  let range = selection_range(selection)
  let start_node = range_node(range, true)
  let end_node = range_node(range, false)
  if !self.host.contains(start_node) || !self.host.contains(end_node) {
    return
  }
  let start_page = node_element(start_node).bind(node => {
    node.closest(".word-page")
  })
  let end_page = node_element(end_node).bind(node => node.closest(".word-page"))
  if start_page is None || end_page is None {
    return
  }
  let anchors : Array[WordAnchor] = []
  let by_key : Map[String, WordAnchor] = Map([])
  let pages : Array[Double] = []
  for fragment in all(self.host, "svg text") {
    if !range_intersects(range, fragment.as_node()) {
      continue
    }
    let contents = text(fragment)
    let length = contents.length()
    let start = if fragment.contains(start_node) {
      range_prefix_length(
        fragment.as_node(),
        start_node,
        range_offset(range, true),
        0,
      )
    } else {
      0
    }
    let end = if fragment.contains(end_node) {
      range_prefix_length(
        fragment.as_node(),
        end_node,
        range_offset(range, false),
        length,
      )
    } else {
      length
    }
    let start = Int32::max(0, Int32::min(length, start))
    let end = Int32::max(0, Int32::min(length, end))
    if end <= start {
      continue
    }
    let page = fragment
      .closest(".word-page")
      .bind(node => {
        if attr(node, "data-page") == "" {
          None
        } else {
          Some(numeric_attr(node, "data-page"))
        }
      })
    if page is Some(page) && !pages.contains(page) {
      pages.push(page)
    }
    let source_id = attr(fragment, "data-source-id")
    let source_part = attr(fragment, "data-source-part")
    let source_path = attr(fragment, "data-source-path")
    let selector = attr(fragment, "data-source-selector")
    let source_start = numeric_attr(fragment, "data-source-start")
    let source_end = numeric_attr(fragment, "data-source-end")
    let logical = selector != "" &&
      fragment.has_attribute("data-source-start") &&
      fragment.has_attribute("data-source-end") &&
      finite(source_start) &&
      finite(source_end) &&
      source_end >= source_start &&
      source_end - source_start == length.to_double()
    let logical_start = if logical {
      source_start + start.to_double()
    } else {
      start.to_double()
    }
    let logical_end = if logical {
      source_start + end.to_double()
    } else {
      end.to_double()
    }
    let rect = fragment.get_bounding_client_rect()
    let key = source_id +
      "|" +
      source_part +
      "|" +
      source_path +
      "|" +
      selector +
      "|" +
      page.map(value => value.to_string()).unwrap_or("") +
      "|" +
      rect.get_left().to_string() +
      "|" +
      rect.get_top().to_string()
    if by_key.get(key) is Some(previous) {
      let old_start = match previous.fields.get("start_utf16") {
        Some(Number(n, ..)) => n
        _ => logical_start
      }
      let old_end = match previous.fields.get("end_utf16") {
        Some(Number(n, ..)) => n
        _ => logical_end
      }
      previous.fields["start_utf16"] = number(
        Double::min(old_start, logical_start),
      )
      previous.fields["end_utf16"] = number(Double::max(old_end, logical_end))
      continue
    }
    let fields : Map[String, Json] = {
      "stability": Json::string(
        if logical {
          let stability = attr(fragment, "data-source-stability")
          if stability == "" {
            "snapshot-relative"
          } else {
            stability
          }
        } else {
          "physical-only"
        },
      ),
      "start_utf16": number(logical_start),
      "end_utf16": number(logical_end),
      // Bounds were clamped to this fragment above. Preserve the browser's
      // UTF-16 slice semantics rather than adjusting Range offsets.
      "text": Json::string(contents.unsafe_substring(start~, end~)),
      "rect": Json::object({
        "left": number(rect.get_left()),
        "top": number(rect.get_top()),
        "width": number(rect.get_width()),
        "height": number(rect.get_height()),
      }),
    }
    if source_id != "" {
      fields["source_id"] = Json::string(source_id)
    }
    if source_part != "" {
      fields["part"] = Json::string(source_part)
    }
    if page is Some(page) {
      fields["page"] = number(page)
    }
    if logical {
      fields["path"] = Json::string(source_path)
      fields["selector"] = Json::string(selector)
      fields["actionable"] = Json::boolean(
        attr(fragment, "data-source-actionable") == "true",
      )
      let status = attr(fragment, "data-source-anchor-status")
      fields["paragraph_anchor_status"] = Json::string(
        if status == "" {
          "missing"
        } else {
          status
        },
      )
      for
        (key, attribute) in [
          ("reason", "data-source-reason"),
          ("para_id", "data-source-para-id"),
        ] {
        let value = attr(fragment, attribute)
        if value != "" {
          fields[key] = Json::string(value)
        }
      }
      let runs = attr(fragment, "data-source-runs")
        .split(",")
        .filter(value => value != "")
        .map(value => Json::string(value.to_owned()))
        .collect()
      if !runs.is_empty() {
        fields["runs"] = Json::array(runs)
      }
    }
    let anchor : WordAnchor = {
      fields,
      page: page.unwrap_or(0.0),
      top: rect.get_top(),
      left: rect.get_left(),
    }
    by_key[key] = anchor
    anchors.push(anchor)
  }
  anchors.stable_sort_by((a, b) => {
    let page = compare_top(a.page, b.page)
    if page != 0 {
      return page
    }
    let top = compare_top(a.top, b.top)
    if top != 0 {
      top
    } else {
      compare_top(a.left, b.left)
    }
  })
  pages.sort_by(compare_top)
  self.emit(
    Json::object({
      "schema": Json::string("office.selection.v1"),
      "file": Json::string(self.document_file),
      "format": Json::string(if self.kind == "word" { "docx" } else { "xlsx" }),
      "kind": Json::string("word"),
      "text": Json::string(selected_text),
      "anchors": Json::array(anchors.map(anchor => Json::object(anchor.fields))),
      "pages": Json::array(pages.map(number)),
    }),
  )
}

///|
fn SurfaceRuntime::cancel_pointer(self : SurfaceRuntime) -> Unit {
  if self.pointer is Some(ExcelPointer(pointer)) {
    clear_cells(pointer.table)
  }
  self.pointer = None
}

///|
fn SurfaceRuntime::install_selection(self : SurfaceRuntime) -> Unit {
  self.listen(self.host.as_event_target(), "pointerdown", event => {
    if !self.alive {
      return
    }
    if event.to_mouse_event() is Some(mouse) {
      if mouse.get_button() != 0 {
        return
      }
    }
    guard event.target().to_element() is Some(target) else { return }
    if target.closest(".office-context-rail, .office-note-marker") is Some(_) {
      return
    }
    if self.kind == "excel" {
      guard target.closest("td") is Some(cell) && attr(cell, "data-cell") != "" else {
        return
      }
      guard cell.closest("table") is Some(table) else { return }
      self.clear_context()
      event.prevent_default()
      self.pointer = Some(ExcelPointer({ anchor: cell, active: cell, table, }))
      ignore(self.select_cells(cell, cell))
    } else if target.closest(".word-page") is Some(_) {
      self.clear_context()
      self.pointer = Some(WordPointer)
    }
  })
  self.listen(@dom.window().as_event_target(), "pointermove", event => {
    if !self.alive {
      return
    }
    guard self.pointer is Some(ExcelPointer(pointer)) else { return }
    guard event.to_mouse_event() is Some(mouse) else { return }
    guard @dom.document()
      .element_from_point(mouse.get_client_x(), mouse.get_client_y())
      .bind(node => node.closest("td"))
      is Some(cell) &&
      attr(cell, "data-cell") != "" &&
      cell_bounds(cell) is Some(_) &&
      cell.closest("table") is Some(table) &&
      same(table, pointer.table) else {
      return
    }
    if same(cell, pointer.active) {
      return
    }
    pointer.active = cell
    ignore(self.select_cells(pointer.anchor, pointer.active))
  })
  self.listen(@dom.window().as_event_target(), "pointerup", event => {
    if !self.alive {
      return
    }
    guard self.pointer is Some(pointer) else { return }
    if event
      .target()
      .to_element()
      .bind(target => {
        target.closest(".office-context-rail, .office-note-marker")
      })
      is Some(_) {
      self.pointer = None
      return
    }
    match pointer {
      ExcelPointer(pointer) => self.excel_selection(pointer)
      WordPointer => self.word_selection()
    }
    self.pointer = None
  })
  self.listen(@dom.window().as_event_target(), "pointercancel", _ => {
    self.cancel_pointer()
  })
  self.listen(@dom.document().as_event_target(), "selectionchange", _ => {
    if self.pointer is None &&
      !self.annotation_selection() &&
      self.kind == "word" {
      self.word_selection()
    }
  })
  if self.kind == "excel" {
    self.listen(self.host.as_event_target(), "dragstart", event => {
      event.prevent_default()
    })
  }
}
