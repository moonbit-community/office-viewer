///|
/// A screen preview uses the same minimum-device-pixel stroke adjustment as
/// PDF viewers. Preserve the authored width in the SVG attribute so exports
/// and future zoom levels can always recover it.
fn adjust_presentation_strokes(host : @dom.Element) -> Unit {
  let ratio = display_pixel_ratio()
  for path in all(host, ".pptx-slide svg path[stroke-width]") {
    if attr(path, "fill") != "none" || attr(path, "vector-effect") != "" {
      continue
    }
    let halves : Array[String] = attr(path, "d")
      .split("L")
      .map(part => part.to_owned())
      .collect()
    guard halves is [start, end] && start.has_prefix("M") else { continue }
    fn point(text : String) -> Array[Double] {
      text
      .trim()
      .split(" ")
      .filter(value => !value.is_empty())
      .map(value => browser_number(value.to_owned()))
      .collect()
    }
    let first = point(start[1:].to_owned())
    let last = point(end)
    guard first is [x1, y1] && last is [x2, y2] else { continue }
    guard svg_screen_matrix(path).to_option() is Some(matrix) else { continue }
    let b : Double = matrix.get_with_string("b")
    let c : Double = matrix.get_with_string("c")
    let a : Double = matrix.get_with_string("a")
    let d : Double = matrix.get_with_string("d")
    let dx = x2 - x1
    let dy = y2 - y1
    let length = (dx * dx + dy * dy).sqrt()
    let screen_x = a * dx + c * dy
    let screen_y = b * dx + d * dy
    let screen_length = (screen_x * screen_x + screen_y * screen_y).sqrt()
    guard length > 0 && screen_length > 0 else { continue }
    // Perpendicular stroke scale under rotation, flips and nonuniform groups.
    let scale = (a * d - b * c).abs() * length / screen_length * ratio
    let authored = numeric_attr(path, "stroke-width")
    guard scale > 0 && authored > 0 && !scale.is_inf() && !scale.is_nan() else {
      continue
    }
    let adjusted = if authored * scale < 1 { 1.0 / scale } else { authored }
    style(path, "stroke-width", adjusted.to_string() + "px")
    if screen_x.abs() < 1.0e-8 || screen_y.abs() < 1.0e-8 {
      style(path, "shape-rendering", "crispEdges")
    }
  }
}

///|
fn prepare_presentation_strokes(host : @dom.Element) -> Unit {
  call_hook(host, "__officeViewerStrokesCleanup")
  if host.query_selector(".pptx-slide") is None {
    return
  }
  adjust_presentation_strokes(host)
  let stop = observe_resize(host, () => adjust_presentation_strokes(host))
  let target = @dom.window().as_event_target()
  let resize = fn(_ : @dom.Event) { adjust_presentation_strokes(host) }
  target.add_event_listener_with_options("resize", resize, capture=false)
  set_callback(host, "__officeViewerStrokesCleanup", () => {
    stop()
    host.remove_property("__officeViewerStrokesCleanup")
    target.remove_event_listener_with_options("resize", resize, capture=false)
  })
}
