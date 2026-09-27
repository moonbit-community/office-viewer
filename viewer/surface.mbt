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
    @html.div(class="surface-empty", [
      @html.div([
        @html.span(class="brand-mark", "Office Viewer"),
        @html.span("\nNo document loaded."),
      ]),
    ])
  } else {
    let class_name = match config.kind {
      Some(Word) => "surface word-surface word-mode"
      Some(Excel) => "surface sheet-surface sheet-mode"
      None => "surface"
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
        selection => scheduler.add(on_selection(selection)),
      )
    },
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
