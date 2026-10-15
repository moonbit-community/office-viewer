///|
pub fn set_surface_zoom(id : String, zoom : Double) -> Unit {
  if @dom.document().get_element_by_id(id).to_option() is Some(host) {
    style(host, "--office-zoom", zoom.to_string())
    call_hook(host, "__officeLayoutAnnotations")
  }
}

///|
pub fn set_system_font_mode(id : String, enabled : Bool) -> Unit {
  if @dom.document().get_element_by_id(id).to_option() is Some(host) {
    toggle_class(host, "prefer-system-fonts", enabled)
    host.set_attribute(
      "data-office-font-mode",
      if enabled {
        "system"
      } else {
        "bundled"
      },
    )
  }
}

///|
pub fn prepare_font_runtime(id : String) -> Unit {
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) else {
    return
  }
  let faces : Array[String] = []
  for
    family in [
      "Carlito", "Liberation Sans", "Liberation Serif", "Liberation Mono",
    ] {
    for font_style in ["normal", "italic"] {
      for weight in ["400", "700"] {
        faces.push(font_style + " " + weight + " 12px \"" + family + "\"")
      }
    }
  }
  faces.push("normal 400 12px \"Noto Sans SC\"")
  load_fonts(FixedArray::from_array(faces), success => {
    host.set_attribute(
      "data-office-fonts",
      if success {
        "bundled"
      } else {
        "fallback"
      },
    )
  })
}

///|
pub fn apply_search(id : String, query : String) -> Unit {
  guard @dom.document().get_element_by_id(id).to_option() is Some(host) else {
    return
  }
  for node in all(host, ".office-search-hit") {
    toggle_class(node, "office-search-hit", false)
  }
  let needle = locale_lowercase(query.trim().to_owned())
  if needle == "" {
    return
  }
  for node in all(host, "svg text, td") {
    if locale_lowercase(text(node)).contains(needle) {
      toggle_class(node, "office-search-hit", true)
    }
  }
}
