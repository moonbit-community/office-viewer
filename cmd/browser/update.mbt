///|
fn surface_kind(model : Model) -> String {
  match model.kind {
    Some(Word) => "word"
    Some(Excel) => "excel"
    None => ""
  }
}

///|
fn surface_command(kind : String, emit : @rabbita.Emit[Msg]) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    scheduler => {
      install_surface_interactions("document-surface", kind, selection => {
        scheduler.add(emit(SelectionChanged(selection)))
      })
    },
    kind=@cmd.after_render,
  )
}

///|
fn visual_command(model : Model) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => {
      set_surface_zoom("document-surface", model.zoom.to_double() / 100.0)
      prepare_font_runtime("document-surface")
      set_system_font_mode("document-surface", model.prefer_system_fonts)
      apply_search("document-surface", model.query)
    },
    kind=@cmd.after_render,
  )
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
                  surface_command(surface_kind(next), emit),
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
                  surface_command(surface_kind(next), emit),
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
