///|
fn compare_top(left : Double, right : Double) -> Int32 {
  if left < right {
    -1
  } else if left > right {
    1
  } else {
    0
  }
}

///|
fn SurfaceRuntime::position_annotation_markers(
  self : SurfaceRuntime,
  rail : @dom.Element,
) -> Unit {
  let groups = self.markers.copy()
  groups.stable_sort_by((a, b) => {
    compare_top(
      a.targets[0].get_bounding_client_rect().get_top(),
      b.targets[0].get_bounding_client_rect().get_top(),
    )
  })
  let mut previous_top = -28.0
  for group in groups {
    let target = self.visible_target(group.targets)
    set_boolean(group.button, "hidden", target is None)
    if target is Some(target) && self.kind == "word" {
      let source_top = Double::max(
        0.0,
        target.get_bounding_client_rect().get_top() -
        rail.get_bounding_client_rect().get_top(),
      )
      let top = Double::max(source_top, previous_top + 28.0)
      px(group.button, "top", top)
      px(group.button, "--office-marker-offset", source_top - top)
      previous_top = top
    }
  }
}

///|
fn SurfaceRuntime::position_annotations(self : SurfaceRuntime) -> Unit {
  if !self.alive {
    return
  }
  guard self.rail is Some(rail) && rail.get_is_connected() else { return }
  guard self.placed is Some(placed) &&
    self.unplaced is Some(unplaced) &&
    self.unplaced_heading is Some(unplaced_heading) else {
    return
  }
  self.position_annotation_markers(rail)
  if !get_boolean(rail, "open") {
    return
  }
  let origin = placed.get_bounding_client_rect().get_top() +
    rail.get_scroll_top()
  let ordered = self.cards.map(record => {
    (record, self.visible_target(record.context.targets))
  })
  ordered.stable_sort_by((a, b) => {
    let a_top = a.1
      .map(node => node.get_bounding_client_rect().get_top())
      .unwrap_or(@double.infinity)
    let b_top = b.1
      .map(node => node.get_bounding_client_rect().get_top())
      .unwrap_or(@double.infinity)
    compare_top(a_top, b_top)
  })
  let mut bottom = 0.0
  let mut unplaced_count = 0
  for (record, target) in ordered {
    let destination = if target is Some(_) { placed } else { unplaced }
    if !record.card
      .get_parent_element()
      .map(parent => same(parent, destination))
      .unwrap_or(false) {
      append(destination, record.card)
    }
    toggle_class(record.card, "office-context-unplaced-card", target is None)
    match target {
      Some(target) => {
        let source_top = Double::max(
          0.0,
          target.get_bounding_client_rect().get_top() - origin,
        )
        let top = Double::max(source_top, bottom)
        px(record.card, "top", top)
        record.card.set_attribute("data-source-top", source_top.to_string())
        px(record.card, "--office-card-offset", source_top - top)
        let card_height = height(record.card)
        let card_height = if card_height != 0.0 {
          card_height
        } else {
          record.card.get_bounding_client_rect().get_height()
        }
        bottom = top +
          (if card_height != 0.0 { card_height } else { 1.0 }) +
          12.0
        set_text(
          record.hint,
          source_label(record.context.item) +
          (if record.context.item.id != record.root.id {
            " · Reply location"
          } else {
            ""
          }),
        )
      }
      None => {
        unplaced_count += 1
        style(record.card, "top", "")
        let anchor = record.context.item.anchor
        let reason = if string(field(anchor, "sheet")) !=
          attr(self.host, "data-office-active-sheet") {
          "Another worksheet"
        } else if string(field(anchor, "preview")) == "outside-preview" {
          "Outside the preview"
        } else {
          "Outside the current view"
        }
        set_text(
          record.hint,
          if self.kind == "word" {
            "Location unavailable"
          } else {
            source_label(record.context.item) + " · " + reason
          },
        )
      }
    }
  }
  px(placed, "height", bottom)
  set_boolean(unplaced_heading, "hidden", unplaced_count == 0)
  set_text(
    unplaced_heading,
    (if self.kind == "word" { "Other comments" } else { "Other notes" }) +
    " · " +
    unplaced_count.to_string(),
  )
}
