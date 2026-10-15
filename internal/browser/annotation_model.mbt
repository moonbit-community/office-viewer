///|
fn SurfaceRuntime::read_annotations(self : SurfaceRuntime) -> Unit {
  let envelope = metadata(self.host, ".office-annotations-metadata")
  if string(field(envelope, "schema")) != "office.annotations.v1" {
    return
  }
  for data in array(field(envelope, "items")) {
    let identity = field(data, "document")
    guard field(identity, "file") is String(file) && file == self.document_file else {
      continue
    }
    guard field(identity, "sha256") is String(sha256) else { continue }
    let anchor = field(data, "anchor")
    let targets = self.resolve_annotation_targets(anchor)
    let body = if string(field(data, "text_status")) == "available" {
      match field(data, "text") {
        String(body) => Some(body)
        _ => None
      }
    } else {
      None
    }
    self.annotations.push({
      data,
      id: string(field(data, "id")),
      file,
      sha256,
      parent_id: string(field(data, "parent_id")),
      body,
      anchor,
      targets,
    })
  }
  self.visible_count = Int32::min(200, self.annotations.length())
}

///|
fn SurfaceRuntime::resolve_annotation_targets(
  self : SurfaceRuntime,
  anchor : Json,
) -> Array[@dom.Element] {
  if self.kind == "excel" && string(field(anchor, "kind")) == "xlsx" {
    let sheet = string(field(anchor, "sheet"))
    let cell = string(field(anchor, "cell"))
    return all(self.host, "td[data-cell]").filter(node => {
      attr(node, "data-sheet") == sheet && attr(node, "data-cell") == cell
    })
  }
  if self.kind == "word" && string(field(anchor, "kind")) == "docx" {
    let locations = array(field(anchor, "locations"))
    return all(
      self.host,
      "svg text[data-comment-path], svg text[data-source-path]",
    ).filter(node => {
      let part = source_attr(node, "part")
      let path = source_attr(node, "path")
      locations.any(location => {
        string(field(location, "part")) == part &&
        string(field(location, "path")) == path
      })
    })
  }
  []
}

///|
fn source_attr(node : @dom.Element, suffix : String) -> String {
  let comment = attr(node, "data-comment-" + suffix)
  if comment != "" {
    comment
  } else {
    attr(node, "data-source-" + suffix)
  }
}

///|
fn SurfaceRuntime::current_annotation(
  self : SurfaceRuntime,
  request : Json,
) -> Annotation? {
  if !self.alive || !self.host.get_is_connected() {
    return None
  }
  let identity = field(request, "document")
  guard field(identity, "file") is String(file) else { return None }
  guard field(identity, "sha256") is String(sha256) else { return None }
  let id = string(field(request, "id"))
  let matches = self.annotations.filter(item => {
    item.id == id && item.file == file && item.sha256 == sha256
  })
  match matches {
    [item] => Some(item)
    _ => None
  }
}

///|
fn SurfaceRuntime::thread_root(
  self : SurfaceRuntime,
  item : Annotation,
) -> Annotation {
  let seen : Map[String, Bool] = Map([])
  let mut current = item
  while current.parent_id != "" {
    if seen.contains(current.id) {
      return item
    }
    seen[current.id] = true
    let parents = self.annotations.filter(candidate => {
      candidate.id == current.parent_id
    })
    guard parents is [parent] && parent.sha256 == item.sha256 else { break }
    current = parent
  }
  current
}

///|
fn SurfaceRuntime::annotation_context(
  self : SurfaceRuntime,
  item : Annotation,
) -> AnnotationContext {
  if !item.targets.is_empty() {
    return { item, targets: item.targets, }
  }
  let root = self.thread_root(item)
  { item: root, targets: if root.id != item.id { root.targets } else { [] }, }
}

///|
fn SurfaceRuntime::visible_target(
  self : SurfaceRuntime,
  targets : Array[@dom.Element],
) -> @dom.Element? {
  targets
  .iter()
  .find_first(node => {
    if !node.get_is_connected() {
      return false
    }
    if node.closest("section") is Some(section) {
      if get_boolean(section, "hidden") {
        return false
      }
    }
    if self.kind != "excel" {
      return true
    }
    let viewport = node
      .closest(".office-sheet-viewport")
      .otherwise(() => node.closest(".office-excel-workbook"))
    guard viewport is Some(viewport) else { return true }
    let r = node.get_bounding_client_rect()
    let v = viewport.get_bounding_client_rect()
    v.get_width() == 0.0 ||
    v.get_height() == 0.0 ||
    (
      r.get_bottom() > v.get_top() &&
      r.get_top() < v.get_bottom() &&
      r.get_right() > v.get_left() &&
      r.get_left() < v.get_right()
    )
  })
}

///|
fn source_label(item : Annotation) -> String {
  if string(field(item.anchor, "kind")) == "xlsx" {
    let sheet = string(field(item.anchor, "sheet"))
    let cell = string(field(item.anchor, "cell"))
    return (if sheet == "" { "Unknown sheet" } else { sheet }) +
      "!" +
      (if cell == "" { "Unknown cell" } else { cell })
  }
  let pages : Array[String] = []
  for node in item.targets {
    if node.closest(".word-page") is Some(page) {
      let value = attr(page, "data-page")
      if value != "" && !pages.contains(value) {
        pages.push(value)
      }
    }
  }
  if pages.is_empty() {
    "Location unavailable"
  } else {
    (if pages.length() > 1 { "Pages " } else { "Page " }) + pages.join(", ")
  }
}

///|
fn SurfaceRuntime::emit_annotation(
  self : SurfaceRuntime,
  item : Annotation,
  start : Int32,
  end : Int32,
  scope : String,
) -> Bool {
  if self.current_annotation(item.data) is None {
    return false
  }
  guard item.body is Some(body) else { return false }
  if start < 0 ||
    end < start ||
    end > body.length() ||
    body.get_view(start~, end~) is None {
    return false
  }
  self.selected_annotation = Some(item.id)
  for entry in all(self.host, ".office-context-entry") {
    toggle_class(
      entry,
      "office-annotation-selected",
      attr(entry, "data-office-annotation-id") == item.id,
    )
  }
  (self.on_annotation_selection)(item.data.stringify(), start, end, scope)
  true
}

///|
fn SurfaceRuntime::select_annotation(
  self : SurfaceRuntime,
  request : Json,
) -> Bool {
  guard self.current_annotation(request) is Some(item) &&
    item.body is Some(body) else {
    return false
  }
  if native_selection().to_option() is Some(selection) {
    selection_clear(selection)
  }
  self.emit_annotation(item, 0, body.length(), "whole")
}

///|
fn SurfaceRuntime::focus_annotation(
  self : SurfaceRuntime,
  request : Json,
) -> Bool {
  match self.current_annotation(request) {
    Some(item) => self.show_annotation(item, navigate=true)
    None => false
  }
}

///|
fn SurfaceRuntime::show_annotation(
  self : SurfaceRuntime,
  item : Annotation,
  navigate~ : Bool = false,
) -> Bool {
  if self.current_annotation(item.data) is None {
    return false
  }
  guard self.rail is Some(rail) else { return false }
  set_boolean(rail, "open", true)
  let context = self.annotation_context(item)
  if context.targets.is_empty() {
    return false
  }
  let sheet = string(field(context.item.anchor, "sheet"))
  if self.kind == "excel" &&
    attr(self.host, "data-office-active-sheet") != sheet {
    self.set_excel_sheet(sheet)
  }
  self.clear_context()
  for target in context.targets {
    toggle_class(target, "office-annotation-focus", true)
  }
  if self.cards
    .iter()
    .find_first(card => card.items.any(entry => entry.id == item.id))
    is Some(record) {
    toggle_class(record.card, "office-context-active", true)
  }
  if navigate {
    context.targets[0].scroll_into_view_with_options(
      block="center",
      inline="center",
    )
  }
  self.schedule_layout()
  true
}
