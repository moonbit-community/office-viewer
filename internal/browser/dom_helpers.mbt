///|
fn element(
  tag : String,
  class_name~ : String = "",
  text~ : String = "",
) -> @dom.Element {
  let node = @dom.document().create_element(tag)
  node.set_attribute("class", class_name)
  node_set_text(node.as_node(), text)
  node
}

///|
fn append(parent : @dom.Element, child : @dom.Element) -> Unit {
  parent.append_child(child.as_node())
}

///|
fn insert_before(
  parent : @dom.Element,
  child : @dom.Element,
  before : @dom.Element,
) -> Unit {
  parent.insert_before(
    child.as_node(),
    @js.Nullable::from_option(Some(before.as_node())),
  )
}

///|
fn remove(node : @dom.Element) -> Unit {
  if node.get_parent_node().to_option() is Some(parent) {
    parent.remove_child(node.as_node())
  }
}

///|
fn text(node : @dom.Element) -> String {
  node_text(node.as_node())
}

///|
fn set_text(node : @dom.Element, value : String) -> Unit {
  node_set_text(node.as_node(), value)
}

///|
fn attr(node : @dom.Element, name : String) -> String {
  node.get_attribute(name).unwrap_or("")
}

///|
fn has_class(node : @dom.Element, name : String) -> Bool {
  attr(node, "class").split(" ").any(part => part == name)
}

///|
fn toggle_class(node : @dom.Element, name : String, enabled : Bool) -> Unit {
  if has_class(node, name) == enabled {
    return
  }
  let names = attr(node, "class")
    .split(" ")
    .filter(part => part != "" && part != name)
    .map(part => part.to_owned())
    .collect()
  if enabled {
    names.push(name)
  }
  node.set_attribute("class", names.join(" "))
}

///|
fn style(node : @dom.Element, name : String, value : String) -> Unit {
  element_style(node).set_property(name, value)
}

///|
fn px(node : @dom.Element, name : String, value : Double) -> Unit {
  style(node, name, value.to_string() + "px")
}

///|
fn width(node : @dom.Element) -> Double {
  node.to_html_element().map(el => el.get_offset_width()).unwrap_or(0.0)
}

///|
fn height(node : @dom.Element) -> Double {
  node.to_html_element().map(el => el.get_offset_height()).unwrap_or(0.0)
}

///|
fn same(left : @dom.Element, right : @dom.Element) -> Bool {
  left.is_same_node(right.as_node())
}

///|
fn node_element(node : @dom.Node) -> @dom.Element? {
  if node.get_node_type() == 1 {
    node.to_element()
  } else {
    node.get_parent_element()
  }
}

///|
fn all(node : @dom.Element, selector : String) -> Array[@dom.Element] {
  node.query_selector_all(selector).iter().collect()
}

///|
fn call_hook(node : @dom.Element, name : String) -> Unit {
  if get_callback(node, name).to_option() is Some(callback) {
    callback()
  }
}

///|
fn field(value : Json, key : String) -> Json {
  match value {
    Object(fields) => fields.get(key).unwrap_or(Json::null())
    _ => Json::null()
  }
}

///|
fn string(value : Json) -> String {
  match value {
    String(value) => value
    _ => ""
  }
}

///|
fn array(value : Json) -> Array[Json] {
  match value {
    Array(values) => values
    _ => []
  }
}

///|
fn number(value : Double) -> Json {
  Json::number(value)
}

///|
fn parse_json(value : String) -> Json {
  @json.parse(value) catch {
    _ => Json::null()
  }
}

///|
fn metadata(host : @dom.Element, selector : String) -> Json {
  host
  .query_selector(selector)
  .map(node => parse_json(text(node)))
  .unwrap_or(Json::null())
}

///|
fn finite(value : Double) -> Bool {
  !value.is_nan() && !value.is_inf()
}

///|
fn[T] Option::otherwise(value : T?, fallback : () -> T?) -> T? {
  match value {
    Some(_) => value
    None => fallback()
  }
}

///|
fn[T] Array::stable_sort_by(
  items : Array[T],
  compare : (T, T) -> Int32,
) -> Unit {
  let indexed : Array[(Int32, T)] = []
  for index, item in items {
    indexed.push((index, item))
  }
  indexed.sort_by((a, b) => {
    let order = compare(a.1, b.1)
    if order == 0 {
      a.0 - b.0
    } else {
      order
    }
  })
  for index, pair in indexed {
    items[index] = pair.1
  }
}

///|
fn numeric_attr(node : @dom.Element, name : String) -> Double {
  browser_number(attr(node, name))
}

///|
fn column_name(column : Int32) -> String {
  for value = column, name = ""; value > 0; {
    let remainder = (value - 1) % 26
    continue (value - 1) / 26,
      (65 + remainder).unsafe_to_char().to_string() + name
  } nobreak {
    name
  }
}

///|
fn svg_icon(path : String, size : String) -> @dom.Element {
  let svg = @dom.document().create_element_ns(@dom.namespace_svg, "svg")
  for
    (key, value) in [
      ("viewBox", "0 0 24 24"),
      ("width", size),
      ("height", size),
      ("fill", "none"),
      ("stroke", "currentColor"),
      ("stroke-width", "1.7"),
      ("aria-hidden", "true"),
    ] {
    svg.set_attribute(key, value)
  }
  let shape = @dom.document().create_element_ns(@dom.namespace_svg, "path")
  shape.set_attribute("d", path)
  append(svg, shape)
  svg
}

///|
fn observe_resize(node : @dom.Element, callback : () -> Unit) -> () -> Unit {
  if !has_resize_observer() {
    return () => ()
  }
  let observer = @dom.ResizeObserver::new((_, _) => callback())
  observer.observe(node)
  () => observer.disconnect()
}
