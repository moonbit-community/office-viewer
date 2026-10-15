///|
fn surface_command(model : Model, emit : @rabbita.Emit[Msg]) -> @rabbita.Cmd {
  @rabbita.batch([
    @viewer.interactions_command(surface_config(model), selection => {
      emit(SelectionChanged(selection))
    }),
    @viewer.pinch_zoom_command(
      surface_config(model),
      zoom => emit(ZoomChanged(zoom.to_string())),
      min_zoom=10,
      max_zoom=300,
    ),
  ])
}

///|
fn visual_command(model : Model, emit : @rabbita.Emit[Msg]) -> @rabbita.Cmd {
  @rabbita.batch([
    @viewer.visual_command(surface_config(model)),
    @viewer.fit_width_command(
      "document-surface",
      surface_config(model).kind,
      zoom => emit(WidthFitted(model.generation, zoom)),
      enabled=model.fit_width,
    ),
  ])
}

///|
fn surface_config(model : Model) -> @viewer.SurfaceConfig {
  {
    surface_id: "document-surface",
    kind: model.kind.map(kind => {
      match kind {
        Word => @viewer.Word
        Excel => @viewer.Excel
        Pdf => @viewer.Pdf
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
fn selection_output(model : Model) -> String {
  if model.selection == "" {
    return ""
  }
  match
    @viewer.materialize_selection(model.selection, {
      mode: model.selection_mode,
      max_preview_utf16: @viewer.default_selection_preview_utf16,
      max_source_anchors: @viewer.default_selection_source_anchors,
    }) {
    Ok(value) => value
    Err(error) => error
  }
}

///|
fn file_ready(
  model : Model,
  generation : Int32,
  payload : FilePayload?,
  emit : @rabbita.Emit[Msg],
) -> (Model, @rabbita.Cmd) {
  if generation != model.generation {
    return (model, @rabbita.none)
  }
  let payload = match payload {
    Some(value) => value
    None =>
      return (
        { ..model, loading: false, status: "File could not be read.", },
        @rabbita.none,
      )
  }
  match parse_document(payload) {
    Err(error) => ({ ..model, loading: false, status: error, }, @rabbita.none)
    Ok(document) => {
      let pdf = document.kind == Pdf
      let next = {
        ..model,
        kind: Some(document.kind),
        fit_width: document.kind == Word,
        zoom: 100,
        file_name: document.name,
        document_file: document.file,
        surface_html: document.surface_html,
        item_count: document.item_count,
        warning: document.warning,
        annotations: document.annotations,
        comments_open: if pdf {
          false
        } else {
          model.comments_open
        },
        status: if pdf {
          "Loading PDF..."
        } else {
          "Ready."
        },
        loading: pdf,
      }
      let command = if pdf {
        @viewer.load_pdf_command(
          surface_config(next),
          payload.bytes,
          selection => emit(PdfSelectionChanged(generation, selection)),
          state => emit(PdfStateChanged(generation, state)),
        )
      } else {
        @rabbita.batch([surface_command(next, emit), visual_command(next, emit)])
      }
      (next, command)
    }
  }
}

///|
fn pdf_state(model : Model, json : String) -> Model {
  let event = @json.parse(json) catch { _ => return model }
  match event {
    { "status": String("ready"), "pages": Number(pages, ..), .. } =>
      {
        ..model,
        item_count: pages.to_int(),
        loading: false,
        password_required: false,
        password: "",
        status: "Ready.",
      }
    { "status": String("page"), "page": Number(page, ..), .. } =>
      { ..model, page: page.to_int(), }
    { "status": String("scale"), "zoom": Number(zoom, ..), .. } =>
      { ..model, zoom: zoom.to_int(), }
    { "status": String("rotation"), "rotation": Number(rotation, ..), .. } =>
      { ..model, rotation: rotation.to_int(), }
    {
      "status": String("search"),
      "current": Number(current, ..),
      "total": Number(total, ..),
      ..
    } =>
      {
        ..model,
        search_current: current.to_int(),
        search_total: total.to_int(),
      }
    { "status": String("password"), "incorrect": incorrect, .. } =>
      {
        ..model,
        password_required: true,
        password: "",
        password_incorrect: incorrect == Json::boolean(true),
        status: "Password required.",
      }
    { "status": String("error"), "message": String(message), .. } =>
      {
        ..model,
        loading: false,
        password_required: false,
        password: "",
        status: message,
      }
    { "status": String("selection-error"), "message": String(message), .. } =>
      { ..model, status: message, }
    _ => model
  }
}

///|
fn begin_read(model : Model) -> Model {
  {
    ..initial_model(),
    zoom: model.zoom,
    inspector_open: model.inspector_open,
    comments_open: model.comments_open,
    annotations_copy_request: model.annotations_copy_request + 1,
    selection_mode: model.selection_mode,
    prefer_system_fonts: model.prefer_system_fonts,
    generation: model.generation + 1,
    loading: true,
    status: "Reading file...",
  }
}

///|
fn capture_selection(model : Model, selection : String) -> Model {
  {
    ..model,
    selection,
    resolved_selection: "",
    resolving: false,
    status: if is_annotation_selection(selection) {
      "Annotation text selected. It is ready to send to an Agent."
    } else {
      "Selection captured."
    },
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
    FileSelected => {
      let next = begin_read(model)
      (
        next,
        @rabbita.batch([
          @viewer.dispose_interactions_command("document-surface"),
          @rabbita.perform(
            payload => emit(FileReady(next.generation, payload)),
            async fn() noraise { read_selected_file("file-input") },
          ),
        ]),
      )
    }
    FileReady(generation, payload) | DropReady(generation, payload) =>
      file_ready(model, generation, payload, emit)
    DropRequested(event) => {
      let next = begin_read(model)
      (
        next,
        @rabbita.batch([
          @viewer.dispose_interactions_command("document-surface"),
          @rabbita.perform(
            payload => emit(DropReady(next.generation, payload)),
            async fn() noraise { read_dropped_file(event) },
          ),
        ]),
      )
    }
    ZoomChanged(raw) => {
      let zoom = @string.parse_int(raw) catch { _ => model.zoom }
      let next = {
        ..model,
        zoom: zoom.clamp(min=10, max=300),
        fit_width: false,
      }
      (next, visual_command(next, emit))
    }
    WidthFitted(generation, zoom) =>
      if generation == model.generation && model.fit_width {
        ({ ..model, zoom, }, @rabbita.none)
      } else {
        (model, @rabbita.none)
      }
    SearchChanged(query) => {
      let next = { ..model, query, search_current: 0, search_total: 0, }
      (next, visual_command(next, emit))
    }
    SelectionChanged(selection) =>
      (capture_selection(model, selection), @rabbita.none)
    PdfSelectionChanged(generation, selection) =>
      if generation == model.generation && model.kind == Some(Pdf) {
        (capture_selection(model, selection), @rabbita.none)
      } else {
        (model, @rabbita.none)
      }
    PdfStateChanged(generation, state) =>
      if generation == model.generation && model.kind == Some(Pdf) {
        (pdf_state(model, state), @rabbita.none)
      } else {
        (model, @rabbita.none)
      }
    SelectionModeChanged(mode) =>
      ({ ..model, selection_mode: mode, }, @rabbita.none)
    ResolveSelection => {
      if model.kind != Some(Pdf) || model.selection == "" {
        return (model, @rabbita.none)
      }
      let reference = selection_output(model)
      (
        {
          ..model,
          resolving: true,
          resolved_selection: "",
          status: "Resolving PDF selection...",
        },
        @rabbita.perform(
          result => {
            emit(SelectionResolved(model.generation, model.selection, result))
          },
          async fn() noraise {
            @viewer.resolve_pdf_selection("document-surface", reference)
          },
        ),
      )
    }
    SelectionResolved(generation, selection, result) => {
      if generation != model.generation || selection != model.selection {
        return (model, @rabbita.none)
      }
      match result {
        Ok(value) =>
          (
            {
              ..model,
              resolving: false,
              resolved_selection: value,
              status: "Selection resolved from PDF text items.",
            },
            @rabbita.none,
          )
        Err(error) =>
          ({ ..model, resolving: false, status: error, }, @rabbita.none)
      }
    }
    PageChanged(raw) => {
      let page = @string.parse_int(raw) catch { _ => model.page }
      (model, @viewer.pdf_page_command("document-surface", page))
    }
    RotatePdf =>
      (
        model,
        @viewer.pdf_rotation_command(
          "document-surface",
          (model.rotation + 90) % 360,
        ),
      )
    FitWidth =>
      if model.kind == Some(Pdf) {
        (model, @viewer.pdf_fit_width_command("document-surface"))
      } else {
        let next = { ..model, fit_width: true, }
        (next, visual_command(next, emit))
      }
    FindPdf(previous) =>
      (model, @viewer.pdf_find_command("document-surface", previous))
    PasswordChanged(password) => ({ ..model, password, }, @rabbita.none)
    SubmitPassword =>
      (
        {
          ..model,
          password: "",
          password_required: false,
          status: "Unlocking PDF...",
        },
        @viewer.pdf_password_command("document-surface", model.password),
      )
    CancelPassword =>
      (
        {
          ..initial_model(),
          generation: model.generation + 1,
          annotations_copy_request: model.annotations_copy_request + 1,
          status: "PDF loading cancelled.",
        },
        @viewer.dispose_interactions_command("document-surface"),
      )
    ToggleSystemFonts => {
      let next = { ..model, prefer_system_fonts: !model.prefer_system_fonts, }
      (next, visual_command(next, emit))
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
    SelectAnnotation(annotation) =>
      (
        model,
        @viewer.select_annotation_command("document-surface", annotation, on_result=selected => {
          emit(AnnotationSelected(selected))
        }),
      )
    AnnotationSelected(selected) =>
      if selected {
        (model, @rabbita.none)
      } else {
        (
          {
            ..model,
            status: "Comment text unavailable or no longer belongs to the open document.",
          },
          @rabbita.none,
        )
      }
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
      copy_to_clipboard(selection_output(model))
      ({ ..model, status: "Selection copied.", }, @rabbita.none)
    }
    ClearSelection =>
      (
        {
          ..model,
          selection: "",
          resolved_selection: "",
          resolving: false,
          status: "Selection cleared.",
        },
        @viewer.clear_selection_command("document-surface"),
      )
  }
}
