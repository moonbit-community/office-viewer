///|
fn surface_command(model : Model, emit : @rabbita.Emit[Msg]) -> @rabbita.Cmd {
  @viewer.interactions_command(surface_config(model), selection => {
    emit(SelectionChanged(selection))
  })
}

///|
fn visual_command(model : Model) -> @rabbita.Cmd {
  @viewer.visual_command(surface_config(model))
}

///|
fn surface_config(model : Model) -> @viewer.SurfaceConfig {
  {
    surface_id: "document-surface",
    kind: model.kind.map(kind => {
      match kind {
        Word => @viewer.Word
        Excel => @viewer.Excel
      }
    }),
    document_file: model.document_file,
    surface_html: model.surface_html,
    zoom: model.zoom,
    query: model.query,
    prefer_system_fonts: model.prefer_system_fonts,
  }
}

///|
fn update(
  model : Model,
  message : Msg,
  emit : @rabbita.Emit[Msg],
) -> (Model, @rabbita.Cmd) {
  match message {
    OpenRequested =>
      (
        model,
        @cmd.custom_cmd(
          _ => open_file_picker("file-input"),
          kind=@cmd.after_render,
        ),
      )
    FileSelected =>
      (
        { ..model, loading: true, status: "Reading file...", },
        @rabbita.perform(payload => emit(FileReady(payload)), async fn() noraise {
          read_selected_file("file-input")
        }),
      )
    FileReady(payload) =>
      match payload {
        None =>
          (
            { ..model, loading: false, status: "Choose a Word or Excel file.", },
            @rabbita.none,
          )
        Some(payload) =>
          match parse_document(payload) {
            Ok(document) => {
              let next = {
                ..model,
                kind: Some(document.kind),
                file_name: document.name,
                document_file: document.file,
                surface_html: document.surface_html,
                item_count: document.item_count,
                warning: document.warning,
                status: "Ready. Drag across text or cells to create an Agent selection.",
                selection: "",
                loading: false,
              }
              (
                next,
                @rabbita.batch([
                  surface_command(next, emit),
                  visual_command(next),
                ]),
              )
            }
            Err(error) =>
              ({ ..model, loading: false, status: error, }, @rabbita.none)
          }
      }
    DropRequested(event) =>
      (
        { ..model, loading: true, status: "Reading dropped file...", },
        @rabbita.perform(payload => emit(DropReady(payload)), async fn() noraise {
          read_dropped_file(event)
        }),
      )
    DropReady(payload) =>
      match payload {
        None =>
          (
            {
              ..model,
              status: "The dropped item is not a supported office file.",
            },
            @rabbita.none,
          )
        Some(payload) =>
          match parse_document(payload) {
            Ok(document) => {
              let next = {
                ..model,
                kind: Some(document.kind),
                file_name: document.name,
                document_file: document.file,
                surface_html: document.surface_html,
                item_count: document.item_count,
                warning: document.warning,
                status: "Ready. Drag across text or cells to create an Agent selection.",
                selection: "",
                loading: false,
              }
              (
                next,
                @rabbita.batch([
                  surface_command(next, emit),
                  visual_command(next),
                ]),
              )
            }
            Err(error) =>
              ({ ..model, loading: false, status: error, }, @rabbita.none)
          }
      }
    ZoomChanged(raw) => {
      let zoom = @string.parse_int(raw) catch { _ => model.zoom }
      let bounded = if zoom < 50 { 50 } else if zoom > 200 { 200 } else { zoom }
      let next = { ..model, zoom: bounded, }
      (next, visual_command(next))
    }
    SearchChanged(query) => {
      let next = { ..model, query, }
      (next, visual_command(next))
    }
    SelectionChanged(selection) =>
      (
        {
          ..model,
          selection,
          status: "Selection captured. It is ready to send to an Agent.",
        },
        @rabbita.none,
      )
    ToggleSystemFonts => {
      let enabled = !model.prefer_system_fonts
      let next = {
        ..model,
        prefer_system_fonts: enabled,
        status: if enabled {
          "Installed font preview enabled. Page layout stays fixed."
        } else {
          "Bundled font preview enabled."
        },
      }
      (next, visual_command(next))
    }
    ToggleInspector =>
      ({ ..model, inspector_open: !model.inspector_open, }, @rabbita.none)
    CopySelection => {
      copy_to_clipboard(model.selection)
      (
        { ..model, status: "Selection copied to the clipboard.", },
        @rabbita.none,
      )
    }
    ClearSelection =>
      ({ ..model, selection: "", status: "Selection cleared.", }, @rabbita.none)
  }
}
