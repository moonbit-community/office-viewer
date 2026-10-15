///|
/// Keep DOCX pages or the active worksheet within the available surface width.
/// Store the reported percentage in the host state. Pass enabled=false after
/// manual zoom; call again with enabled=true when the user chooses Fit width.
/// Surface disposal retires the resize observer and pending animation frame.
pub fn fit_width_command(
  config : SurfaceConfig,
  on_zoom : (Int32) -> @cmd.Cmd,
  enabled~ : Bool = true,
) -> @cmd.Cmd {
  @cmd.custom_cmd(
    scheduler => {
      @browser_runtime.configure_fit_width(
        config.surface_id,
        surface_kind(config.kind),
        enabled,
        zoom => scheduler.add(on_zoom(zoom)),
      )
    },
    kind=@cmd.after_render,
  )
}
