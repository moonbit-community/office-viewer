///|
fn author_name(data : Json) -> String {
  let author = normalize_name(string(field(data, "author")))
  if author == "" {
    "Unknown author"
  } else {
    author
  }
}

///|
fn avatar_label(author : String, initials : String) -> String {
  let words : Array[String] = []
  let word = StringBuilder()
  for ch in author {
    if ch.is_whitespace() {
      if !word.is_empty() {
        words.push(word.to_string())
        word.reset()
      }
    } else {
      word.write_char(ch)
    }
  }
  if !word.is_empty() {
    words.push(word.to_string())
  }
  let compact = StringBuilder()
  for ch in initials {
    if !ch.is_whitespace() {
      compact.write_char(ch)
    }
  }
  let value = if initials != "" {
    compact.to_string()
  } else if words.length() > 1 {
    let first = browser_graphemes(words[0])
    let last = browser_graphemes(words[words.length() - 1])
    return uppercase(first[0] + last[0])
  } else {
    author
  }
  let characters = browser_graphemes(value)
  uppercase(characters.iter().take(2).collect().join(""))
}

///|
pub fn annotation_avatar(data : Json) -> @dom.Element {
  let author = normalize_name(string(field(data, "author")))
  let initials = normalize_name(string(field(data, "initials")))
  let label = avatar_label(author, initials)
  let avatar = element("span", class_name="office-annotation-avatar")
  avatar.set_attribute("aria-hidden", "true")
  if label != "" {
    let palette = [
      ("#e0e7ff", "#3730a3"),
      ("#dcfce7", "#166534"),
      ("#fce7f3", "#9d174d"),
      ("#fef3c7", "#92400e"),
      ("#e0f2fe", "#075985"),
      ("#f3e8ff", "#6b21a8"),
    ]
    let mut hash = 0U
    let color_key = if author == "" { initials } else { author }
    for ch in color_key {
      hash = hash * 31U + ch.to_int().reinterpret_as_uint()
    }
    let (background, color) = palette[(hash %
      palette.length().reinterpret_as_uint()).reinterpret_as_int()]
    style(avatar, "--office-avatar-background", background)
    style(avatar, "--office-avatar-color", color)
    set_text(avatar, label)
  } else {
    let icon = svg_icon("M5 20v-1a7 7 0 0 1 14 0v1", "18")
    icon.set_attribute("stroke-linecap", "round")
    let head = @dom.document().create_element_ns(@dom.namespace_svg, "circle")
    for (key, value) in [("cx", "12"), ("cy", "8"), ("r", "3")] {
      head.set_attribute(key, value)
    }
    icon.insert_before(
      head.as_node(),
      @js.Nullable::from_option(
        icon.query_selector("path").map(node => node.as_node()),
      ),
    )
    append(avatar, icon)
  }
  avatar
}
