///|
/// Install trackpad pinch/Ctrl+wheel zoom without intercepting ordinary scroll.
/// Preserve the reported percentage in the host's SurfaceConfig. Disposal
/// removes the listener and cancels any pending zoom frame.
pub fn pinch_zoom_command(
  config : SurfaceConfig,
  on_zoom : (Int32) -> @cmd.Cmd,
  min_zoom~ : Int32 = 25,
  max_zoom~ : Int32 = 400,
) -> @cmd.Cmd {
  @cmd.custom_cmd(
    scheduler => {
      @browser_runtime.install_pinch_zoom(
        config.surface_id,
        surface_kind(config.kind),
        min_zoom,
        max_zoom,
        zoom => scheduler.add(on_zoom(zoom)),
      )
    },
    kind=@cmd.after_render,
  )
}
