///|
priv struct WidthFit {
  host : @dom.Element
  kind : String
  mut on_zoom : (Int32) -> Unit
  mut alive : Bool
  mut frame : Double?
  mut last_zoom : Int32?
  cleanups : Array[() -> Unit]
}

///|
let width_fits : Map[String, WidthFit] = Map([])

///|
fn WidthFit::schedule(self : WidthFit) -> Unit {
  if !self.alive || self.frame is Some(_) {
    return
  }
  self.frame = Some(
    @dom.window().request_animation_frame(_ => {
      self.frame = None
      if self.alive && self.host.get_is_connected() {
        self.apply()
      }
    }),
  )
}

///|
fn WidthFit::apply(self : WidthFit) -> Unit {
  let host = self.host
  let current = browser_number(computed_style(host, "--office-zoom"))
  let current = if current > 0 { current } else { 1.0 }
  let pages = if self.kind == "word" {
    all(host, ".word-page")
  } else {
    all(host, "section:not([hidden]) table")
  }
  let natural = pages.fold(init=0.0, (widest, page) => {
    let page_width = page.get_bounding_client_rect().get_width() / current
    if page_width > widest {
      page_width
    } else {
      widest
    }
  })
  let mut available = host.get_client_width() -
    css_pixels(computed_style(host, "padding-left")) -
    css_pixels(computed_style(host, "padding-right"))
  if host.query_selector(".office-context-rail") is Some(rail) {
    available -= rail.get_bounding_client_rect().get_width()
    if host.query_selector(".office-context-layout") is Some(layout) {
      available -= css_pixels(computed_style(layout, "column-gap"))
    }
  }
  if available <= 0 || natural <= 0 {
    return
  }
  let zoom = (available / natural * 100.0)
    .floor()
    .to_int()
    .clamp(min=10, max=100)
  if self.last_zoom != Some(zoom) ||
    (current * 100.0 - zoom.to_double()).abs() > 0.5 {
    self.last_zoom = Some(zoom)
    set_surface_zoom(attr(host, "id"), zoom.to_double() / 100.0)
    (self.on_zoom)(zoom)
  }
}

///|
/// Follow the width allocated by a host, including its annotation gutter.
/// Excel fitting is explicit; hosts normally keep large workbooks at 100%.
pub fn configure_fit_width(
  id : String,
  kind : String,
  enabled : Bool,
  on_zoom : (Int32) -> Unit,
) -> Unit {
  if !enabled {
    dispose_fit_width(id)
    return
  }
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) &&
    (kind == "word" || kind == "excel") else {
    return
  }
  if width_fits.get(id) is Some(fit) && fit.kind == kind && same(fit.host, host) {
    fit.on_zoom = on_zoom
    fit.schedule()
    return
  }
  dispose_fit_width(id)
  let fit : WidthFit = {
    host,
    kind,
    on_zoom,
    alive: true,
    frame: None,
    last_zoom: None,
    cleanups: [],
  }
  width_fits[id] = fit
  fit.cleanups.push(observe_resize(host, () => fit.schedule()))
  if host.query_selector(".office-context-rail") is Some(rail) {
    fit.cleanups.push(observe_resize(rail, () => fit.schedule()))
  }
  fonts_ready(() => fit.schedule())
  fit.schedule()
}

///|
pub fn dispose_fit_width(id : String) -> Unit {
  if width_fits.get(id) is Some(fit) {
    fit.alive = false
    if fit.frame is Some(frame) {
      @dom.window().cancel_animation_frame(frame)
    }
    for cleanup in fit.cleanups {
      cleanup()
    }
    width_fits.remove(id)
  }
}
