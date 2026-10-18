///|
priv struct Annotation {
  data : Json
  id : String
  file : String
  sha256 : String
  parent_id : String
  body : String?
  anchor : Json
  targets : Array[@dom.Element]
}

///|
priv struct AnnotationContext {
  item : Annotation
  targets : Array[@dom.Element]
}

///|
priv struct AnnotationCard {
  card : @dom.Element
  root : Annotation
  items : Array[Annotation]
  context : AnnotationContext
  hint : @dom.Element
}

///|
priv struct MarkerGroup {
  button : @dom.Element
  targets : Array[@dom.Element]
  items : Array[Annotation]
  page : @dom.Element?
  part : String
  path : String
}

///|
priv struct CellPointer {
  anchor : @dom.Element
  mut active : @dom.Element
  table : @dom.Element
  anchor_bounds : CellBounds
}

///|
priv enum SurfacePointer {
  WordPointer
  ExcelPointer(CellPointer)
}

///|
priv struct SurfaceRuntime {
  id : String
  host : @dom.Element
  kind : String
  document_file : String
  on_selection : (String) -> Unit
  on_annotation_selection : (String, Int32, Int32, String) -> Unit
  cleanups : Array[() -> Unit]
  annotations : Array[Annotation]
  cards : Array[AnnotationCard]
  markers : Array[MarkerGroup]
  bodies : Array[(@dom.Element, Annotation)]
  lazy_sheets : Map[String, LazySheet]
  mut lazy_styles : Json
  mut lazy_query : String
  mut alive : Bool
  mut pointer : SurfacePointer?
  mut word_selection_payload : String?
  mut selected_annotation : String?
  mut visible_count : Int32
  mut rail : @dom.Element?
  mut placed : @dom.Element?
  mut unplaced : @dom.Element?
  mut unplaced_heading : @dom.Element?
  mut layout_frame : Double?
}

///|
let surfaces : Map[String, SurfaceRuntime] = Map([])

///|
fn SurfaceRuntime::listen(
  self : SurfaceRuntime,
  target : @dom.EventTarget,
  name : String,
  handler : (@dom.Event) -> Unit,
  capture~ : Bool = false,
) -> Unit {
  target.add_event_listener_with_options(name, handler, capture~)
  self.cleanups.push(() => {
    target.remove_event_listener_with_options(name, handler, capture~)
  })
}

///|
fn SurfaceRuntime::emit(self : SurfaceRuntime, value : Json) -> Unit {
  self.clear_context()
  (self.on_selection)(value.stringify(indent=2))
}

///|
fn SurfaceRuntime::clear_context(self : SurfaceRuntime) -> Unit {
  for node in all(self.host, ".office-annotation-focus, .office-context-active") {
    toggle_class(node, "office-annotation-focus", false)
    toggle_class(node, "office-context-active", false)
  }
}

///|
fn SurfaceRuntime::schedule_layout(self : SurfaceRuntime) -> Unit {
  if !self.alive || self.layout_frame is Some(_) {
    return
  }
  self.layout_frame = Some(
    @dom.window().request_animation_frame(_ => {
      self.layout_frame = None
      self.position_annotations()
    }),
  )
}

///|
fn SurfaceRuntime::dispose(self : SurfaceRuntime) -> Unit {
  if !self.alive {
    return
  }
  self.alive = false
  call_hook(self.host, "__officeViewerPageThumbnailsCleanup")
  dispose_fit_width(self.id)
  call_hook(self.host, "__officeViewerZoomCleanup")
  self.pointer = None
  for sheet in self.lazy_sheets.values() {
    if sheet.frame is Some(frame) {
      @dom.window().cancel_animation_frame(frame)
    }
  }
  self.lazy_sheets.clear()
  if self.layout_frame is Some(frame) {
    @dom.window().cancel_animation_frame(frame)
  }
  self.layout_frame = None
  self.clear_context()
  for cleanup in self.cleanups {
    cleanup()
  }
  self.cleanups.clear()
  self.cards.clear()
  self.markers.clear()
  self.bodies.clear()
  for
    name in [
      "__officeViewerCleanup", "__officeFocusAnnotation", "__officeSelectAnnotation",
      "__officeLayoutAnnotations", "__officeLayoutObjects", "__officeRevealExcelCell",
    ] {
    self.host.remove_property(name)
  }
  surfaces.remove(self.id)
}

///|
pub fn install_surface_interactions(
  id : String,
  kind : String,
  document_file : String,
  surface_html : String,
  on_selection : (String) -> Unit,
  on_annotation_selection : (String, Int32, Int32, String) -> Unit,
) -> Unit {
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) else {
    return
  }
  call_hook(host, "__officeViewerCleanup")
  // Hosts may replace the root node while retaining its id. Retire the
  // registered snapshot even when its callback is on the detached old root.
  if surfaces.get(id) is Some(previous) {
    if !same(previous.host, host) {
      call_hook(previous.host, "__officeViewerCleanup")
    }
    previous.dispose()
  }
  host.set_inner_html(surface_html)
  let runtime : SurfaceRuntime = {
    id,
    host,
    kind,
    document_file,
    on_selection,
    on_annotation_selection,
    cleanups: [],
    annotations: [],
    cards: [],
    markers: [],
    bodies: [],
    lazy_sheets: Map([]),
    lazy_styles: Json::null(),
    lazy_query: "",
    alive: true,
    pointer: None,
    word_selection_payload: None,
    selected_annotation: None,
    visible_count: 0,
    rail: None,
    placed: None,
    unplaced: None,
    unplaced_heading: None,
    layout_frame: None,
  }
  surfaces[id] = runtime
  set_callback(host, "__officeViewerCleanup", () => runtime.dispose())
  set_callback(host, "__officeLayoutObjects", () => {
    runtime.layout_excel_objects()
  })
  runtime.prepare_lazy_excel()
  runtime.prepare_excel()
  excel_reveal_hook(host, (sheet, row, column) => {
    runtime.reveal_lazy_cell(sheet, row, column)
  })
  runtime.read_annotations()
  runtime.mount_annotations()
  runtime.install_selection()
  set_callback(host, "__officeLayoutAnnotations", () => {
    runtime.schedule_layout()
  })
  annotation_hook(host, "__officeFocusAnnotation", request => {
    runtime.focus_annotation(parse_json(request))
  })
  annotation_hook(host, "__officeSelectAnnotation", request => {
    runtime.select_annotation(parse_json(request))
  })
  runtime.listen(@dom.window().as_event_target(), "keydown", event => {
    if event.to_keyboard_event() is Some(keyboard) {
      if keyboard.key() == "Escape" {
        runtime.clear_context()
      }
    }
  })
  runtime.listen(@dom.window().as_event_target(), "resize", _ => {
    runtime.layout_excel_objects()
    runtime.schedule_layout()
  })
  runtime.listen(
    @dom.document().as_event_target(),
    "scroll",
    event => {
      let inside = match (runtime.rail, event.target().to_node()) {
        (Some(rail), Some(node)) => rail.contains(node)
        (_, _) => false
      }
      if !inside {
        runtime.schedule_layout()
      }
    },
    capture=true,
  )
  runtime.cleanups.push(observe_resize(host, () => runtime.schedule_layout()))
  if runtime.rail is Some(rail) {
    runtime.cleanups.push(observe_resize(rail, () => runtime.schedule_layout()))
  }
  fonts_ready(() => {
    if runtime.alive {
      runtime.layout_excel_objects()
      runtime.schedule_layout()
    }
  })
  runtime.schedule_layout()
}

///|
pub fn dispose_surface_interactions(id : String) -> Unit {
  dispose_fit_width(id)
  if @dom.document().get_element_by_id(id).to_option() is Some(host) {
    call_hook(host, "__officeViewerPageThumbnailsCleanup")
    call_hook(host, "__officeViewerZoomCleanup")
    call_hook(host, "__officeViewerCleanup")
  }
  if surfaces.get(id) is Some(runtime) {
    call_hook(runtime.host, "__officeViewerCleanup")
    runtime.dispose()
  }
}

///|
pub fn focus_surface_annotation(id : String, request : String) -> Bool {
  match surfaces.get(id) {
    Some(runtime) => runtime.focus_annotation(parse_json(request))
    None => false
  }
}

///|
pub fn select_surface_annotation(id : String, request : String) -> Bool {
  match surfaces.get(id) {
    Some(runtime) => runtime.select_annotation(parse_json(request))
    None => false
  }
}

///|
pub fn clear_surface_selection(id : String) -> Unit {
  if surfaces.get(id) is Some(runtime) {
    runtime.cancel_pointer()
  }
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) else {
    return
  }
  if native_selection().to_option() is Some(selection) {
    if selection_anchor(selection).to_option() is Some(anchor) {
      if host.contains(anchor) {
        selection_clear(selection)
      }
    }
  }
  for cell in all(host, ".office-cell-selected, .office-cell-active") {
    toggle_class(cell, "office-cell-selected", false)
    toggle_class(cell, "office-cell-active", false)
  }
}
