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
                annotations: document.annotations,
                copying_annotations: false,
                annotations_copy_request: model.annotations_copy_request + 1,
                annotations_copy_status: "",
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
                annotations: document.annotations,
                copying_annotations: false,
                annotations_copy_request: model.annotations_copy_request + 1,
                annotations_copy_status: "",
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
      (
        {
          ..model,
          inspector_open: !model.inspector_open,
          comments_open: false,
        },
        @rabbita.none,
      )
    ToggleComments =>
      (
        { ..model, comments_open: !model.comments_open, inspector_open: false, },
        @rabbita.none,
      )
    FocusAnnotation(annotation) =>
      (
        model,
        @viewer.focus_annotation_command("document-surface", annotation, on_result=found => {
          emit(AnnotationFocused(found))
        }),
      )
    AnnotationFocused(found) =>
      (
        {
          ..model,
          status: if !found {
            "Location unavailable in this preview. The comment and source location remain in Comments."
          } else if model.kind is Some(Word) {
            "Paragraph context shown. The exact commented text range is unavailable."
          } else {
            "Note cell shown in the preview."
          },
        },
        @rabbita.none,
      )
    CopyAnnotations => {
      if model.kind is None || model.copying_annotations {
        return (model, @rabbita.none)
      }
      let request = model.annotations_copy_request + 1
      (
        {
          ..model,
          copying_annotations: true,
          annotations_copy_request: request,
          annotations_copy_status: "Copying comments JSON...",
        },
        @cmd.custom_cmd(scheduler => {
          copy_annotations_to_clipboard(model.annotations.stringify(indent=2), copied => {
            scheduler.add(emit(AnnotationsCopied(request, copied)))
          })
        }),
      )
    }
    AnnotationsCopied(request, copied) => {
      // A pending clipboard operation must not update a newly loaded file.
      if request != model.annotations_copy_request {
        return (model, @rabbita.none)
      }
      (
        {
          ..model,
          copying_annotations: false,
          annotations_copy_status: if copied {
            "Comments JSON copied to the clipboard."
          } else {
            "Clipboard unavailable. Select and copy the JSON below."
          },
        },
        @rabbita.none,
      )
    }
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
