///|
fn SurfaceRuntime::mount_annotations(self : SurfaceRuntime) -> Unit {
  if !self.annotations.is_empty() {
    let styles = element(
      "style",
      class_name="office-annotation-avatar-styles",
      text=annotation_styles,
    )
    append(self.host, styles)
    self.cleanups.push(() => remove(styles))
    let content = self.host
      .query_selector(
        if self.kind == "word" {
          ".word-pages"
        } else {
          ".excel-workspace"
        },
      )
      .otherwise(() => {
        if self.kind == "word" {
          self.host.query_selector(".word-page")
        } else {
          None
        }
      })
    if content is Some(content) {
      self.mount_annotation_rail(content)
    }
  }
  self.mount_annotation_markers()
}

///|
fn SurfaceRuntime::mount_annotation_rail(
  self : SurfaceRuntime,
  content : @dom.Element,
) -> Unit {
  let layout = element(
    "div",
    class_name="office-context-layout office-context-" + self.kind,
  )
  let document_column = element("div", class_name="office-context-document")
  let rail = element(
    "details",
    class_name="office-context-rail office-default-annotations",
  )
  let page_width = if self.kind == "word" {
    content
    .query_selector(".word-page")
    .map(page => page.get_bounding_client_rect().get_width())
    .unwrap_or(0.0)
  } else {
    0.0
  }
  let host_width = self.host.get_client_width()
  set_boolean(
    rail,
    "open",
    host_width == 0.0 || host_width >= Double::max(680.0, page_width + 352.0),
  )
  rail.set_attribute("aria-label", "Comments alongside the document")
  insert_before(self.host, layout, content)
  append(layout, document_column)
  append(document_column, content)
  append(layout, rail)
  let heading = element("summary", class_name="office-context-heading")
  let heading_text = (if self.kind == "word" { "Comments" } else { "Notes" }) +
    " · " +
    self.annotations.length().to_string()
  heading.set_attribute("data-count", self.annotations.length().to_string())
  append(
    heading,
    element(
      "span",
      class_name="office-context-heading-label",
      text=heading_text,
    ),
  )
  heading.set_attribute("aria-label", heading_text)
  heading.set_attribute("title", heading_text)
  append(rail, heading)
  self.rail = Some(rail)
  self.listen(rail.as_event_target(), "toggle", _ => self.schedule_layout())
  let placed = element("div", class_name="office-context-placed")
  let unplaced_heading = element(
    "div",
    class_name="office-context-unplaced-heading",
  )
  let unplaced = element("div", class_name="office-context-unplaced")
  for node in [placed, unplaced_heading, unplaced] {
    append(rail, node)
  }
  self.placed = Some(placed)
  self.unplaced = Some(unplaced)
  self.unplaced_heading = Some(unplaced_heading)
  self.cleanups.push(() => {
    if layout.get_parent_element() is Some(parent) {
      if same(parent, self.host) {
        insert_before(self.host, content, layout)
      }
    }
    remove(layout)
  })
  self.render_annotation_cards()
}

///|
fn SurfaceRuntime::render_annotation_cards(self : SurfaceRuntime) -> Unit {
  guard self.rail is Some(rail) && self.placed is Some(placed) else { return }
  for record in self.cards {
    remove(record.card)
  }
  self.cards.clear()
  self.bodies.clear()
  if rail.query_selector(".office-context-more") is Some(more) {
    remove(more)
  }
  let groups : Array[(Annotation, Array[Annotation])] = []
  for index in 0..<self.visible_count {
    let item = self.annotations[index]
    let root = self.thread_root(item)
    match groups.iter().find_first(group => group.0.id == root.id) {
      Some((_, items)) => items.push(item)
      None => groups.push((root, [item]))
    }
  }
  for (root, items) in groups {
    // Preserve source order among replies while placing the root first.
    let ordered = items.filter(item => item.id == root.id)
    ordered.append(items.filter(item => item.id != root.id))
    let context_item = if !root.targets.is_empty() {
      root
    } else {
      ordered
      .iter()
      .find_first(item => !item.targets.is_empty())
      .unwrap_or(root)
    }
    let context : AnnotationContext = {
      item: context_item,
      targets: context_item.targets,
    }
    let card = element("aside", class_name="office-context-card")
    card.set_attribute("data-office-thread-id", root.id)
    card.set_attribute(
      "aria-label",
      if string(field(root.data, "type")) == "xlsx.note" {
        "Cell Note"
      } else {
        "Comment thread"
      },
    )
    let header = element("div", class_name="office-context-header")
    append(
      header,
      element(
        "strong",
        class_name="",
        text=if string(field(root.data, "type")) == "xlsx.note" {
          "Note"
        } else {
          "Comment"
        },
      ),
    )
    let locate = element(
      "button",
      class_name="office-context-select",
      text="Show source",
    )
    locate.set_attribute("type", "button")
    set_boolean(
      locate,
      "disabled",
      context.targets.is_empty() &&
      !self.lazy_sheets.contains(string(field(context_item.anchor, "sheet"))),
    )
    locate.add_event_listener("click", _ => {
      if card.get_is_connected() {
        ignore(self.show_annotation(context_item, navigate=true))
      }
    })
    append(header, locate)
    append(card, header)
    let hint = element(
      "div",
      class_name="office-context-location",
      text=source_label(context_item),
    )
    append(card, hint)
    if self.kind == "word" && !context.targets.is_empty() {
      let source_text = context.targets.map(text).join("").trim().to_owned()
      if source_text != "" {
        let quote = element(
          "blockquote",
          class_name="office-context-quote",
          text=source_text,
        )
        quote.set_attribute("title", source_text)
        append(card, quote)
      }
    }
    for entry in ordered {
      append(card, self.annotation_entry(entry, root))
    }
    append(placed, card)
    self.cards.push({ card, root, items: ordered, context, hint, })
  }
  if self.visible_count < self.annotations.length() {
    let more = element("div", class_name="office-context-more")
    let noun = if self.kind == "word" { "comments" } else { "notes" }
    append(
      more,
      element(
        "p",
        class_name="",
        text=self.visible_count.to_string() +
          " of " +
          self.annotations.length().to_string() +
          " " +
          noun,
      ),
    )
    let remaining = Int32::min(
      200,
      self.annotations.length() - self.visible_count,
    )
    let button = element(
      "button",
      class_name="office-context-select",
      text="Show " + remaining.to_string() + " more",
    )
    button.set_attribute("type", "button")
    button.add_event_listener("click", _ => {
      self.visible_count = Int32::min(
        self.annotations.length(),
        self.visible_count + 200,
      )
      self.render_annotation_cards()
    })
    append(more, button)
    append(rail, more)
  }
  self.position_annotations()
  self.schedule_layout()
}

///|
fn SurfaceRuntime::annotation_entry(
  self : SurfaceRuntime,
  entry : Annotation,
  root : Annotation,
) -> @dom.Element {
  let reply = entry.id != root.id
  let row = element(
    "section",
    class_name="office-context-entry" +
      (if reply { " office-context-reply" } else { "" }),
  )
  row.set_attribute("data-office-annotation-id", entry.id)
  toggle_class(
    row,
    "office-annotation-selected",
    self.selected_annotation == Some(entry.id),
  )
  let person = element("div", class_name="office-context-person")
  append(person, annotation_avatar(entry.data))
  append(
    person,
    element(
      "div",
      class_name="office-context-author",
      text=author_name(entry.data),
    ),
  )
  append(row, person)
  let meta : Array[String] = []
  let date = string(field(entry.data, "date"))
  if date != "" {
    let formatted = locale_date(date)
    if formatted != "" {
      meta.push(formatted)
    }
  }
  if field(entry.data, "resolved") == Json::boolean(true) {
    meta.push("Resolved")
  }
  if !meta.is_empty() {
    append(
      row,
      element("div", class_name="office-context-meta", text=meta.join(" · ")),
    )
  }
  if reply {
    append(row, element("div", class_name="office-context-hint", text="Reply"))
  }
  let body = element(
    "div",
    class_name="office-annotation-body",
    text=entry.body.unwrap_or("Comment text unavailable."),
  )
  body.set_attribute("data-office-annotation-body", entry.id)
  body.set_attribute("tabindex", if entry.body is Some(_) { "0" } else { "-1" })
  body.set_attribute(
    "aria-label",
    (if reply { "Reply" } else { "Comment" }) +
    " text by " +
    author_name(entry.data),
  )
  if entry.body == Some("") {
    body.set_attribute("data-empty", "true")
  }
  if entry.body is Some(_) {
    self.bodies.push((body, entry))
  }
  body.add_event_listener("click", _ => {
    if !body.get_is_connected() {
      return
    }
    let collapsed = native_selection()
      .to_option()
      .map(selection_collapsed)
      .unwrap_or(true)
    if collapsed {
      ignore(self.select_annotation(entry.data))
    }
  })
  append(row, body)
  let select = element(
    "button",
    class_name="office-context-select office-context-quote-action",
    text=if string(field(entry.data, "type")) == "xlsx.note" {
      "Quote note"
    } else {
      "Quote comment"
    },
  )
  select.set_attribute("type", "button")
  set_boolean(select, "disabled", entry.body is None)
  select.add_event_listener("click", _ => {
    if select.get_is_connected() {
      ignore(self.select_annotation(entry.data))
    }
  })
  append(row, select)
  row
}

///|
fn SurfaceRuntime::mount_annotation_markers(self : SurfaceRuntime) -> Unit {
  for item in self.annotations {
    for node in item.targets {
      toggle_class(node, "office-has-annotation", true)
      let ids = array(parse_json(attr(node, "data-office-annotation-ids")))
      if !ids.contains(Json::string(item.id)) {
        ids.push(Json::string(item.id))
      }
      node.set_attribute(
        "data-office-annotation-ids",
        Json::array(ids).stringify(),
      )
      let page = node.closest(".word-page")
      let part = source_attr(node, "part")
      let path = source_attr(node, "path")
      let group = self.markers
        .iter()
        .find_first(group => {
          if self.kind == "excel" {
            return same(group.targets[0], node)
          }
          let same_page = match (group.page, page) {
            (Some(a), Some(b)) => same(a, b)
            (None, None) => true
            (_, _) => false
          }
          same_page && group.part == part && group.path == path
        })
        .unwrap_or_else(() => {
          let button = element(
            "button",
            class_name=if self.kind == "excel" {
              "office-note-marker"
            } else {
              "office-comment-marker office-comment-source-marker"
            },
          )
          button.set_attribute("type", "button")
          if self.kind == "word" {
            let icon = svg_icon(
              "M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z", "14",
            )
            icon.set_attribute("stroke-linejoin", "round")
            append(button, icon)
          }
          let group : MarkerGroup = {
            button,
            targets: [],
            items: [],
            page,
            part,
            path,
          }
          self.markers.push(group)
          if self.kind == "excel" {
            append(node, button)
          } else if self.rail is Some(rail) {
            if rail.query_selector("summary") is Some(summary) {
              append(summary, button)
            }
          }
          self.listen(button.as_event_target(), "click", event => {
            event.prevent_default()
            event.stop_propagation()
            guard group.items.get(0) is Some(selected) else { return }
            if self.show_annotation(selected) {
              if self.kind == "excel" {
                ignore(self.select_annotation(selected.data))
              } else {
                self.position_annotations()
                if self.cards
                  .iter()
                  .find_first(record => {
                    record.items.any(item => item.id == selected.id)
                  })
                  is Some(record) {
                  record.card.scroll_into_view_with_options(
                    block="nearest",
                    inline="nearest",
                  )
                }
              }
            }
          })
          self.cleanups.push(() => remove(button))
          group
        })
      if !group.targets.any(target => same(target, node)) {
        group.targets.push(node)
      }
      if !group.items.any(entry => entry.id == item.id) {
        group.items.push(item)
      }
      let authors : Array[String] = []
      for entry in group.items {
        let author = author_name(self.thread_root(entry).data)
        if !authors.contains(author) {
          authors.push(author)
        }
      }
      let authors = authors.join(", ")
      group.button.set_attribute(
        "aria-label",
        (if self.kind == "excel" { "Note" } else { "Comment" }) +
        " · " +
        source_label(item) +
        " · " +
        authors,
      )
      group.button.set_attribute("title", "Show comment by " + authors)
    }
  }
}
