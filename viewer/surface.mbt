///|
pub(all) struct SurfaceConfig {
  surface_id : String
  kind : DocumentKind?
  document_file : String
  surface_html : String
  zoom : Int
  query : String
  prefer_system_fonts : Bool
} derive(Eq)

///|
pub extend SurfaceConfig with Eq::{equal, not_equal}

///|
/// Render only the document surface. The surrounding toolbar, file picker,
/// and selection inspector remain host-owned so OpenSeek can place this view
/// inside its own layout.
pub fn render_surface(config : SurfaceConfig) -> @html.Html {
  if config.surface_html == "" {
    // Keep the root element and its id mounted even when no Office document
    // is active. Embedded hosts use this stable node to retire listeners when
    // a file tab switches back to the code/Markdown viewer.
    @html.div(
      id=config.surface_id,
      class="surface office-viewer-surface office-viewer-surface-empty moonbit-viewer-surface-hidden",
      [
        @html.div(class="surface-empty", [
          @html.div([
            @html.span(class="brand-mark", "Office Viewer"),
            @html.span("\nNo document loaded."),
          ]),
        ]),
      ],
    )
  } else {
    let class_name = match config.kind {
      Some(Word) => "surface office-viewer-surface word-surface word-mode"
      Some(Excel) => "surface office-viewer-surface sheet-surface sheet-mode"
      None => "surface office-viewer-surface"
    }
    @html.div(
      id=config.surface_id,
      class=class_name,
      attrs=@html.Attrs::build().inner_html(config.surface_html),
      ([] : Array[@html.Html]),
    )
  }
}

///|
/// Install native Word selection and Excel range selection after the host has
/// rendered `render_surface`. The callback receives the versioned selection
/// envelope as JSON text and can forward it to the Agent context layer.
pub fn interactions_command(
  config : SurfaceConfig,
  on_selection : (String) -> @rabbita.Cmd,
) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    scheduler => {
      install_surface_interactions(
        config.surface_id,
        surface_kind(config.kind),
        config.document_file,
        config.surface_html,
        selection => scheduler.add(on_selection(selection)),
      )
    },
    kind=@cmd.after_render,
  )
}

///|
/// Retire the browser listeners owned by one surface. The surface root stays
/// mounted so Rabbita can reuse it across tab switches; this command clears
/// the window-level pointer and selection listeners before the next document
/// is shown.
pub fn dispose_interactions_command(surface_id : String) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => dispose_surface_interactions(surface_id),
    kind=@cmd.after_render,
  )
}

///|
/// Apply zoom, font preference, and search highlighting after a render.
pub fn visual_command(config : SurfaceConfig) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => {
      set_surface_zoom(config.surface_id, config.zoom.to_double() / 100.0)
      prepare_font_runtime(config.surface_id)
      set_system_font_mode(config.surface_id, config.prefer_system_fonts)
      apply_search(config.surface_id, config.query)
    },
    kind=@cmd.after_render,
  )
}

///|
fn surface_kind(kind : DocumentKind?) -> String {
  match kind {
    Some(Word) => "word"
    Some(Excel) => "excel"
    None => ""
  }
}

///|
/// Navigate to an annotation in the currently installed surface. DOCX uses
/// proven paragraph context only, not a claimed exact comment range. False
/// means the cell/paragraph has no matching rendered node (including notes
/// outside preview bounds and unjoined Word documents).
pub fn focus_annotation_command(
  surface_id : String,
  annotation : Json,
  on_result? : (Bool) -> @rabbita.Cmd,
) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    scheduler => {
      let found = focus_surface_annotation(surface_id, annotation.stringify())
      match on_result {
        Some(callback) => scheduler.add(callback(found))
        None => ()
      }
    },
    kind=@cmd.after_render,
  )
}
