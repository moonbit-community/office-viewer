///|
priv struct PinchZoom {
  host : @dom.Element
  kind : String
  on_zoom : (Int32) -> Unit
  mut alive : Bool
  mut frame : Double?
  mut target : Double?
  mut observed : Double?
  mut x : Double
  mut y : Double
}

///|
fn PinchZoom::current(self : PinchZoom) -> Double? {
  if self.kind == "pdf" {
    pdf_current_zoom(self.host).to_option()
  } else {
    let zoom = browser_number(computed_style(self.host, "--office-zoom"))
    Some(if zoom > 0 { zoom * 100.0 } else { 100.0 })
  }
}

///|
fn PinchZoom::apply(self : PinchZoom) -> Unit {
  self.frame = None
  guard self.alive &&
    self.host.get_is_connected() &&
    self.current() is Some(current) &&
    self.target is Some(target) else {
    return
  }
  let zoom = target.round().to_int()
  if (zoom.to_double() - current).abs() < 0.01 {
    return
  }
  if self.kind == "pdf" {
    pdf_zoom_at(self.host, zoom, self.x, self.y)
  } else {
    let sheet = self.host.query_selector(".sheet-document")
    let anchor = if self.kind == "excel" {
      self.host
      .query_selector("section:not([hidden]) table")
      .otherwise(() => sheet)
    } else {
      let pages = all(self.host, ".word-page")
      pages
      .filter(page => {
        let rect = page.get_bounding_client_rect()
        rect.get_top() <= self.y && rect.get_bottom() >= self.y
      })
      .get(0)
      .otherwise(() => pages.get(0))
    }
    let scroller = if self.kind == "excel" {
      self.host
      .query_selector("section:not([hidden]) .office-sheet-viewport")
      .unwrap_or(zoom_scroller(sheet.unwrap_or(self.host)))
    } else {
      zoom_scroller(self.host)
    }
    let before = anchor.map(page => page.get_bounding_client_rect())
    set_surface_zoom(attr(self.host, "id"), zoom.to_double() / 100.0)
    if anchor is Some(anchor) && before is Some(before) {
      let after = anchor.get_bounding_client_rect()
      let ratio = zoom.to_double() / current
      let css_zoom = if self.kind == "excel" {
        zoom.to_double() / 100.0
      } else {
        browser_number(computed_style(scroller, "zoom"))
      }
      let css_zoom = if css_zoom > 0 { css_zoom } else { 1.0 }
      scroller.set_scroll_left(
        scroller.get_scroll_left() +
        (after.get_left() + (self.x - before.get_left()) * ratio - self.x) /
        css_zoom,
      )
      scroller.set_scroll_top(
        scroller.get_scroll_top() +
        (after.get_top() + (self.y - before.get_top()) * ratio - self.y) /
        css_zoom,
      )
    }
  }
  self.observed = Some(zoom.to_double())
  (self.on_zoom)(zoom)
}

///|
fn zoom_scroller(host : @dom.Element) -> @dom.Element {
  for node = Some(host)
      node is Some(element)
      node = element.get_parent_element() {
    let overflow = computed_style(element, "overflow-x") +
      computed_style(element, "overflow-y")
    if overflow.contains("auto") || overflow.contains("scroll") {
      return element
    }
  }
  host
}

///|
pub fn install_pinch_zoom(
  id : String,
  kind : String,
  minimum : Int32,
  maximum : Int32,
  on_zoom : (Int32) -> Unit,
) -> Unit {
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) else {
    return
  }
  call_hook(host, "__officeViewerZoomCleanup")
  let runtime : PinchZoom = {
    host,
    kind,
    on_zoom,
    alive: true,
    frame: None,
    target: None,
    observed: None,
    x: 0.0,
    y: 0.0,
  }
  let wheel = fn(event : @dom.Event) {
    guard event.to_wheel_event() is Some(wheel) &&
      wheel.get_ctrl_key() &&
      runtime.alive &&
      runtime.current() is Some(current) else {
      return
    }
    event.prevent_default()
    event.stop_propagation()
    if runtime.target is None ||
      runtime.observed
      .map(value => (value - current).abs() > 0.01)
      .unwrap_or(true) {
      runtime.target = Some(current)
    }
    let unit = match wheel.get_delta_mode() {
      1 => 16.0
      2 => host.get_client_height()
      _ => 1.0
    }
    let delta = (wheel.get_delta_y() * unit).clamp(min=-100.0, max=100.0)
    runtime.target = Some(
      (runtime.target.unwrap_or(current) * @math.exp(-delta / 100.0)).clamp(
        min=minimum.to_double(),
        max=maximum.to_double(),
      ),
    )
    runtime.observed = Some(current)
    runtime.x = wheel.get_client_x()
    runtime.y = wheel.get_client_y()
    if runtime.frame is None {
      runtime.frame = Some(
        @dom.window().request_animation_frame(_ => runtime.apply()),
      )
    }
  }
  host.add_event_listener_with_options("wheel", wheel, passive=false)
  set_callback(host, "__officeViewerZoomCleanup", () => {
    runtime.alive = false
    host.remove_event_listener("wheel", wheel)
    if runtime.frame is Some(frame) {
      @dom.window().cancel_animation_frame(frame)
    }
    host.remove_property("__officeViewerZoomCleanup")
  })
}
