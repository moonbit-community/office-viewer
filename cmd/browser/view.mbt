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
fn empty_surface() -> Html {
  div(class="surface-empty", [
    div([
      span(class="brand-mark", "Office Viewer"),
      span("\nDrop a .docx or .xlsx file here, or use Open file to start."),
    ]),
  ])
}

///|
fn surface(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let content = if model.surface_html == "" {
    empty_surface()
  } else {
    div(
      id="document-surface",
      class="surface",
      attrs=@html.Attrs::build().inner_html(model.surface_html),
      ([] : Array[Html]),
    )
  }
  let drop_attrs = @html.Attrs::build()
    .on_dragover(event => {
      event.prevent_default()
      @rabbita.none
    })
    .on_drop(event => emit(DropRequested(event)))
  div(id="surface-wrap", class=surface_class(model), attrs=drop_attrs, [content])
}

///|
fn inspector(model : Model, emit : @rabbita.Emit[Msg]) -> Html {
  let panel_class = if model.inspector_open {
    "side-panel open"
  } else {
    "side-panel"
  }
  div(class=panel_class, [
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
          on_click=emit(ToggleInspector),
          if model.inspector_open {
            "Hide selection"
          } else {
            "Selection"
          },
        ),
      ]),
    ]),
    div(class="statusbar", [
      span(class="status-file", "\{kind_label(model.kind)} · \{file_label}"),
      span(class="status-selection", model.status),
      span(
        "\{model.item_count} \{if model.kind is Some(Excel) { "sheets" } else { "pages" }}",
      ),
    ]),
    div(class="viewer-workspace", [surface(model, emit), inspector(model, emit)]),
  ])
}
