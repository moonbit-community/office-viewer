///|
using @html {button, div, h2, input, span}

///|
using @rabbita {type Html}

///|
fn icon(name : String) -> Html {
  span(
    class="office-icon office-icon-" + name,
    attrs=@html.Attrs::build().aria_hidden("true"),
    "",
  )
}

///|
fn tool_button(
  name : String,
  title : String,
  action : @rabbita.Cmd,
  disabled? : Bool = false,
) -> Html {
  button(
    class="button icon",
    title~,
    disabled~,
    attrs=@html.Attrs::build().aria_label(title),
    on_click=action,
    [icon(name)],
  )
}

///|
fn kind_label(kind : DocumentKind?) -> String {
  match kind {
    Some(Word) => "Word"
    Some(Excel) => "Excel"
    Some(Pdf) => "PDF"
    None => ""
  }
}

///|
fn surface_class(model : Model) -> String {
  match model.kind {
    Some(Word) => "surface-wrap word-mode"
    Some(Excel) => "surface-wrap sheet-mode"
    Some(Pdf) => "surface-wrap pdf-mode"
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
    .on_drop(event => {
      event.prevent_default()
      emit(DropRequested(event))
    })
  div(id="surface-wrap", class=surface_class(model), attrs=drop_attrs, [
    @viewer.render_surface(surface_config(model)),
  ])
}

///|
fn mode_button(
  model : Model,
  emit : @rabbita.Emit[Msg],
  mode : @viewer.SelectionMaterializationMode,
  label : String,
) -> Html {
  let selected = model.selection_mode == mode
  button(
    class=if selected { "mode-option active" } else { "mode-option" },
    attrs=@html.Attrs::build().aria_pressed(
      if selected {
        "true"
      } else {
        "false"
      },
    ),
    on_click=emit(SelectionModeChanged(mode)),
    label,
  )
}

///|
fn resolved_text(value : String) -> String {
  if value == "" {
    return ""
  }
  let parsed = @json.parse(value) catch { _ => return value }
  match parsed {
    { "text": String(text), .. } => text
    _ => value
  }
}

///|
fn inspector(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  div(
    id="selection-panel",
    class=if model.inspector_open { "side-panel open" } else { "side-panel" },
    [
      div(class="panel-heading", [
        h2("Selection"),
        tool_button("panel-right", "Close selection", emit(ToggleInspector)),
      ]),
      div(
        class="mode-control",
        attrs=@html.Attrs::build().role("group").aria_label("Selection output"),
        [
          mode_button(model, emit, @viewer.Full, "Full"),
          mode_button(model, emit, @viewer.Preview, "Preview"),
          mode_button(model, emit, @viewer.Deferred, "Deferred"),
        ],
      ),
      div(class="side-panel-actions", [
        tool_button(
          "copy",
          "Copy selection",
          emit(CopySelection),
          disabled=model.selection == "",
        ),
        tool_button(
          "trash-2",
          "Clear selection",
          emit(ClearSelection),
          disabled=model.selection == "",
        ),
        button(
          class="button resolve-button",
          hidden=model.kind != Some(Pdf),
          disabled=model.selection == "" || model.resolving,
          on_click=emit(ResolveSelection),
          [
            icon("scan-text"),
            span(
              if model.resolving {
                "Resolving..."
              } else {
                "Resolve reference"
              },
            ),
          ],
        ),
      ]),
      @html.pre(
        id="selection-output",
        class="selection-content",
        if model.selection == "" {
          "No selection."
        } else {
          selection_output(model)
        },
      ),
      div(class="resolved-content", hidden=model.resolved_selection == "", [
        h2("Resolved Content"),
        @html.pre(
          id="resolved-output",
          class="selection-content",
          resolved_text(model.resolved_selection),
        ),
      ]),
    ],
  )
}

///|
fn pdf_toolbar(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  div(class="pdf-toolbar", hidden=model.kind != Some(Pdf), [
    div(class="pdf-page-control", [
      tool_button(
        "chevron-left",
        "Previous page",
        emit(PageChanged((model.page - 1).to_string())),
        disabled=model.page <= 1 || model.loading,
      ),
      input(
        input_type=Number,
        id="pdf-page",
        class="page-input",
        min=1,
        max=model.item_count,
        value=model.page.to_string(),
        on_change=emit.map(value => PageChanged(value)),
        attrs=@html.Attrs::build().aria_label("Page number"),
      ),
      span(class="page-total", "/ \{model.item_count}"),
      tool_button(
        "chevron-right",
        "Next page",
        emit(PageChanged((model.page + 1).to_string())),
        disabled=model.page >= model.item_count || model.loading,
      ),
    ]),
    div(class="pdf-view-controls", [
      tool_button(
        "wrap-text",
        "Fit page width",
        emit(FitWidth),
        disabled=model.loading,
      ),
      tool_button(
        "rotate-cw",
        "Rotate clockwise",
        emit(RotatePdf),
        disabled=model.loading,
      ),
    ]),
    div(class="pdf-find-controls", hidden=model.query == "", [
      span(
        class="search-count",
        "\{model.search_current} / \{model.search_total}",
      ),
      tool_button(
        "chevron-left",
        "Previous match",
        emit(FindPdf(true)),
        disabled=model.search_total == 0,
      ),
      tool_button(
        "chevron-right",
        "Next match",
        emit(FindPdf(false)),
        disabled=model.search_total == 0,
      ),
    ]),
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
    Some(Json::String(text)) => text
    _ =>
      if annotation_string(annotation, "text_status") == Some("ambiguous") {
        "Comment text unavailable (ambiguous source)."
      } else {
        "Comment text unavailable."
      }
  }
  let has_body = annotation_string(annotation, "text_status") ==
    Some("available") &&
    annotation_field(annotation, "text") is Some(Json::String(_))
  let identity = annotation_field(annotation, "document").unwrap_or(
    Json::null(),
  )
  let body_attrs = if has_body {
    @html.Attrs::build()
    .data_set(
      "office-annotation-body",
      annotation_string(annotation, "id").unwrap_or(""),
    )
    .data_set("office-annotation-surface", "document-surface")
    .data_set(
      "office-annotation-file",
      annotation_string(identity, "file").unwrap_or(""),
    )
    .data_set(
      "office-annotation-sha256",
      annotation_string(identity, "sha256").unwrap_or(""),
    )
    .tabindex(0)
  } else {
    @html.Attrs::build()
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
    div(class="annotation-text", attrs=body_attrs, text),
    div(class="annotation-meta", "\{resolution} · \{parent_label}"),
    div(class="annotation-location", annotation_location(annotation)),
    div(class="annotation-hint", annotation_preview_label(annotation)),
    button(
      class="button annotation-select",
      disabled=!has_body,
      on_click=emit(SelectAnnotation(annotation)),
      if annotation_string(annotation, "type") == Some("xlsx.note") {
        "Select Note"
      } else {
        "Select comment"
      },
    ),
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
fn password_dialog(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  div(class="modal-backdrop", hidden=!model.password_required, [
    @html.form(
      class="password-dialog",
      attrs=@html.Attrs::build()
        .role("dialog")
        .aria_modal("true")
        .aria_label("Unlock PDF")
        .on_submit(event => {
          event.prevent_default()
          emit(SubmitPassword)
        }),
      [
        icon("key-round"),
        h2("Unlock PDF"),
        input(
          input_type=Password,
          id="pdf-password",
          placeholder="Password",
          value=model.password,
          on_input=emit.map(value => PasswordChanged(value)),
          attrs=@html.Attrs::build().aria_label("PDF password"),
        ),
        div(
          class="password-error",
          hidden=!model.password_incorrect,
          "Incorrect password.",
        ),
        div(class="dialog-actions", [
          button(
            class="button",
            on_click=emit(CancelPassword),
            attrs=@html.Attrs::build().type_("button"),
            "Cancel",
          ),
          button(
            class="button primary",
            disabled=model.password == "",
            attrs=@html.Attrs::build().type_("submit"),
            "Unlock",
          ),
        ]),
      ],
    ),
  ])
}

///|
fn view(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let file_label = if model.file_name == "" {
    "No file loaded"
  } else {
    model.file_name
  }
  div(class="viewer-shell", [
    div(class="topbar", [
      div(class="brand", [
        span(class="brand-mark", "OV"),
        div(class="brand-copy", [
          span(class="brand-title", "Office Viewer"),
          span(class="brand-subtitle", "DOCX / XLSX / PDF"),
        ]),
      ]),
      div(class="file-actions", [
        input(
          input_type=File,
          id="file-input",
          accept=".docx,.xlsx,.pdf",
          hidden=true,
          on_change=emit.map(_ => FileSelected),
        ),
        button(class="button primary", on_click=emit(OpenRequested), [
          icon("file-up"),
          span(if model.loading { "Loading..." } else { "Open file" }),
        ]),
        input(
          input_type=Search,
          class="search-box",
          placeholder="Find in document",
          value=model.query,
          on_input=emit.map(value => SearchChanged(value)),
          attrs=@html.Attrs::build().aria_label("Find in document"),
        ),
      ]),
      div(class="view-actions", [
        button(
          class=system_font_button_class(model),
          hidden=model.kind == Some(Pdf),
          title="Prefer installed fonts",
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
          tool_button(
            "zoom-out",
            "Zoom out",
            emit(ZoomChanged((model.zoom - 10).to_string())),
            disabled=model.zoom <= 25,
          ),
          input(
            input_type=Range,
            id="zoom-input",
            min=25,
            max=300,
            step=5,
            value=model.zoom.to_string(),
            on_input=emit.map(value => ZoomChanged(value)),
            attrs=@html.Attrs::build().aria_label("Zoom"),
          ),
          span(class="zoom-value", "\{model.zoom}%"),
          tool_button(
            "zoom-in",
            "Zoom in",
            emit(ZoomChanged((model.zoom + 10).to_string())),
            disabled=model.zoom >= 300,
          ),
        ]),
        button(
          class="button icon",
          title="Selection",
          attrs=@html.Attrs::build()
            .aria_label("Selection")
            .aria_controls("selection-panel")
            .aria_expanded(if model.inspector_open { "true" } else { "false" }),
          on_click=emit(ToggleInspector),
          [icon("panel-right")],
        ),
        button(
          class=if model.comments_open {
            "button comments-toggle active"
          } else {
            "button comments-toggle"
          },
          hidden=model.kind == Some(Pdf),
          attrs=@html.Attrs::build()
            .aria_controls("comments-panel")
            .aria_expanded(if model.comments_open { "true" } else { "false" }),
          on_click=emit(ToggleComments),
          "Comments (\{annotation_count(model)})",
        ),
      ]),
    ]),
    pdf_toolbar(model, emit),
    div(class="statusbar", attrs=@html.Attrs::build().aria_live("polite"), [
      span(class="status-file", "\{kind_label(model.kind)} · \{file_label}"),
      span(class="status-selection", model.status),
      span(
        class="status-comments",
        hidden=model.kind == Some(Pdf),
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
    password_dialog(model, emit),
  ])
}
