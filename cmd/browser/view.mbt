///|
using @html {button, div, h2, input, span}

///|
using @rabbita {type Html}

///|
fn kind_label(kind : DocumentKind?) -> String {
  match kind {
    Some(Word) => "Word"
    Some(Excel) => "Excel"
    None => ""
  }
}

///|
fn surface_class(model : Model) -> String {
  match model.kind {
    Some(Word) => "surface-wrap word-mode"
    Some(Excel) => "surface-wrap sheet-mode"
    None => "surface-wrap"
  }
}

///|
fn surface(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let drop_attrs = @html.Attrs::build()
    .on_dragover(event => {
      event.prevent_default()
      @rabbita.none
    })
    .on_drop(event => emit(DropRequested(event)))
  div(id="surface-wrap", class=surface_class(model), attrs=drop_attrs, [
    @viewer.render_surface(surface_config(model)),
  ])
}

///|
fn inspector(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let panel_class = if model.inspector_open {
    "side-panel open"
  } else {
    "side-panel"
  }
  div(id="selection-panel", class=panel_class, [
    h2("Selection"),
    div(class="side-panel-actions", [
      button(class="button", on_click=emit(CopySelection), "Copy"),
      button(class="button", on_click=emit(ClearSelection), "Clear"),
    ]),
    div(
      class="selection-content",
      if model.selection == "" {
        "No selection yet."
      } else {
        model.selection
      },
    ),
    div(
      class="selection-content",
      if model.warning == "" {
        ""
      } else {
        model.warning
      },
    ),
  ])
}

///|
fn annotation_field(value : Json, key : String) -> Json? {
  match value {
    Json::Object(fields) => fields.get(key)
    _ => None
  }
}

///|
fn annotation_string(value : Json, key : String) -> String? {
  match annotation_field(value, key) {
    Some(Json::String(text)) if text != "" => Some(text)
    _ => None
  }
}

///|
fn annotation_array(value : Json, key : String) -> Array[Json] {
  match annotation_field(value, key) {
    Some(Json::Array(items)) => items
    _ => []
  }
}

///|
fn annotation_count(model : Model) -> Int {
  annotation_array(model.annotations, "items").length()
}

///|
fn annotation_location(annotation : Json) -> String {
  let anchor = annotation_field(annotation, "anchor").unwrap_or(Json::null())
  if annotation_string(anchor, "kind") == Some("xlsx") {
    let sheet = annotation_string(anchor, "sheet").unwrap_or("Unknown sheet")
    let cell = annotation_string(anchor, "cell").unwrap_or("Unknown cell")
    "\{sheet}!\{cell}"
  } else {
    let part = annotation_string(anchor, "part").unwrap_or("Unknown Word part")
    let locations = annotation_array(anchor, "locations")
    let paths = locations.filter_map(location => {
      annotation_string(location, "path")
    })
    if paths.length() > 0 {
      "\{part} · \{paths.join(", ")}"
    } else {
      let sources = annotation_array(annotation, "source_anchors")
      let source_paths = sources.filter_map(source => {
        match annotation_string(source, "start") {
          Some(path) => Some(path)
          None =>
            match annotation_string(source, "end") {
              Some(path) => Some(path)
              None => annotation_string(source, "story")
            }
        }
      })
      if source_paths.is_empty() {
        "No source location recorded."
      } else {
        source_paths.join(", ")
      }
    }
  }
}

///|
fn annotation_preview_label(annotation : Json) -> String {
  let anchor = annotation_field(annotation, "anchor").unwrap_or(Json::null())
  if annotation_string(anchor, "preview") == Some("outside-preview") {
    "Outside preview. This note is retained in the list and JSON."
  } else if annotation_string(anchor, "kind") == Some("docx") {
    "Source location only. Paragraph context may be shown; the commented text range is unavailable."
  } else {
    "Location is shown only when its cell is available in the preview."
  }
}

///|
fn annotation_card(annotation : Json, emit : @rabbita.Emit[Msg]) -> Html {
  let author = annotation_string(annotation, "author").unwrap_or(
    "Unknown author",
  )
  let date = annotation_string(annotation, "date").unwrap_or("Date unknown")
  let initials = annotation_string(annotation, "initials")
  let parent = annotation_string(annotation, "parent_id")
  let parent_source = annotation_string(annotation, "parent_source_id")
  let resolution = match annotation_field(annotation, "resolved") {
    Some(Json::True) => "Resolved"
    Some(Json::False) => "Open"
    _ => "Resolution unknown"
  }
  let parent_label = match parent {
    Some(id) => "Reply to \{id}"
    None =>
      match parent_source {
        Some(id) => "Unlinked parent source: \{id}"
        None => "No parent recorded"
      }
  }
  let text = match annotation_field(annotation, "text") {
    Some(Json::String(text)) =>
      if text == "" {
        "(Empty comment)"
      } else {
        text
      }
    _ =>
      if annotation_string(annotation, "text_status") == Some("ambiguous") {
        "Comment text unavailable (ambiguous source)."
      } else {
        "Comment text unavailable."
      }
  }
  let card_class = if parent is Some(_) {
    "annotation-card annotation-reply"
  } else {
    "annotation-card"
  }
  div(class=card_class, [
    div(class="annotation-heading", [
      span(class="annotation-author", author),
      span(
        class="annotation-kind",
        if annotation_string(annotation, "type") == Some("xlsx.note") {
          "Note"
        } else {
          "Comment"
        },
      ),
    ]),
    div(
      class="annotation-meta",
      match initials {
        Some(initials) => "\{date} · \{initials}"
        None => date
      },
    ),
    div(class="annotation-text", text),
    div(class="annotation-meta", "\{resolution} · \{parent_label}"),
    div(class="annotation-location", annotation_location(annotation)),
    div(class="annotation-hint", annotation_preview_label(annotation)),
    button(
      class="button annotation-locate",
      title=if annotation_string(annotation, "type") == Some("docx.comment") {
        "Show mapped paragraph context without guessing the commented text range"
      } else {
        "Locate only when the source cell is available in the preview"
      },
      on_click=emit(FocusAnnotation(annotation)),
      if annotation_string(annotation, "type") == Some("docx.comment") {
        "Show paragraph context"
      } else {
        "Locate if visible"
      },
    ),
  ])
}

///|
fn comments_panel(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let panel_class = if model.comments_open {
    "side-panel comments-panel open"
  } else {
    "side-panel comments-panel"
  }
  let items = annotation_array(model.annotations, "items")
  let cards : Array[Html] = items.map(item => annotation_card(item, emit))
  let warnings : Array[Html] = annotation_array(model.annotations, "warnings").filter_map(warning => {
      match warning {
        Json::String(text) => Some(div(class="annotation-warning", text))
        _ => None
      }
    },
  )
  div(id="comments-panel", class=panel_class, [
    h2("Comments (\{items.length()})"),
    div(
      class="annotation-panel-hint",
      "Read-only Word comments and Excel Notes.",
    ),
    div(class="annotation-warnings", warnings),
    div(
      class="annotation-empty",
      hidden=items.length() > 0,
      if model.kind is None {
        "Open a Word or Excel file to read comments."
      } else {
        "No supported comments or notes found."
      },
    ),
    div(class="annotation-list", cards),
    @html.details(class="annotation-json-details", hidden=model.kind is None, [
      @html.summary("Structured comments JSON"),
      div(class="annotation-json-actions", [
        button(
          class="button",
          disabled=model.copying_annotations,
          on_click=emit(CopyAnnotations),
          if model.copying_annotations {
            "Copying..."
          } else {
            "Copy JSON"
          },
        ),
      ]),
      div(
        class="annotation-copy-status",
        attrs=@html.Attrs::build().role("status").aria_live("polite"),
        model.annotations_copy_status,
      ),
      @html.pre(
        class="annotation-json",
        attrs=@html.Attrs::build().tabindex(0).aria_label("Comments JSON"),
        model.annotations.stringify(indent=2),
      ),
    ]),
  ])
}

///|
fn system_font_button_class(model : Model) -> String {
  if model.prefer_system_fonts {
    "button icon active"
  } else {
    "button icon"
  }
}

///|
fn view(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let file_label = if model.file_name == "" {
    "No file loaded"
  } else {
    model.file_name
  }
  let loading_label = if model.loading { "Loading..." } else { "Open file" }
  div(class="viewer-shell", [
    div(class="topbar", [
      div(class="brand", [
        span(class="brand-mark", "OV"),
        div(class="brand-copy", [
          span(class="brand-title", "Office Viewer"),
          span(class="brand-subtitle", "read-only high fidelity preview"),
        ]),
      ]),
      div(class="file-actions", [
        input(
          input_type=File,
          id="file-input",
          accept=".docx,.xlsx",
          hidden=true,
          on_change=emit.map(_ => FileSelected),
        ),
        button(
          class="button primary",
          on_click=emit(OpenRequested),
          disabled=model.loading,
          loading_label,
        ),
        input(
          input_type=Search,
          class="search-box",
          placeholder="Find in document",
          value=model.query,
          on_input=emit.map(value => SearchChanged(value)),
        ),
      ]),
      div(class="view-actions", [
        button(
          class=system_font_button_class(model),
          title=if model.prefer_system_fonts {
            "Use bundled fonts"
          } else {
            "Prefer installed fonts"
          },
          attrs=@html.Attrs::build().aria_pressed(
            if model.prefer_system_fonts {
              "true"
            } else {
              "false"
            },
          ),
          on_click=emit(ToggleSystemFonts),
          "Aa",
        ),
        div(class="zoom-control", [
          button(
            class="button icon",
            title="Zoom out",
            on_click=emit(ZoomChanged((model.zoom - 10).to_string())),
            "−",
          ),
          input(
            input_type=Range,
            min=50,
            max=200,
            step=10,
            value=model.zoom.to_string(),
            on_input=emit.map(value => ZoomChanged(value)),
          ),
          span(class="zoom-value", "\{model.zoom}%"),
          button(
            class="button icon",
            title="Zoom in",
            on_click=emit(ZoomChanged((model.zoom + 10).to_string())),
            "+",
          ),
        ]),
        button(
          class="button",
          attrs=@html.Attrs::build()
            .aria_controls("selection-panel")
            .aria_expanded(if model.inspector_open { "true" } else { "false" }),
          on_click=emit(ToggleInspector),
          if model.inspector_open {
            "Hide selection"
          } else {
            "Selection"
          },
        ),
        button(
          class=if model.comments_open {
            "button comments-toggle active"
          } else {
            "button comments-toggle"
          },
          attrs=@html.Attrs::build()
            .aria_controls("comments-panel")
            .aria_expanded(if model.comments_open { "true" } else { "false" }),
          on_click=emit(ToggleComments),
          "Comments (\{annotation_count(model)})",
        ),
      ]),
    ]),
    div(class="statusbar", [
      span(class="status-file", "\{kind_label(model.kind)} · \{file_label}"),
      span(class="status-selection", model.status),
      span(
        class="status-comments",
        "\{annotation_count(model)} comments / notes",
      ),
      span(
        "\{model.item_count} \{if model.kind is Some(Excel) { "sheets" } else { "pages" }}",
      ),
    ]),
    div(class="viewer-workspace", [
      surface(model, emit),
      inspector(model, emit),
      comments_panel(model, emit),
    ]),
  ])
}
