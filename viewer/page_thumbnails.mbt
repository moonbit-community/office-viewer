///|
/// Render a host-neutral thumbnail list. Keep this node mounted beside (not
/// inside) `render_surface`; the runtime owns its children after rendering.
/// Hosts own the surrounding heading, toolbar, visibility and responsive layout.
/// Labels may change on a later render; `{page}` is replaced by the page number.
pub fn render_page_thumbnails(
  id : String,
  page_label~ : String = "Page {page}",
  unavailable_label~ : String = "Preview unavailable",
) -> @html.Html {
  @html.div(
    id~,
    class="office-page-list",
    attrs=@html.Attrs::build()
      .data_set("page-label", page_label)
      .data_set("unavailable-label", unavailable_label),
    ([] : Array[@html.Html]),
  )
}

///|
/// Bind a thumbnail list to the current Word or ready PDF surface. The list
/// renders nearby pages, follows document scrolling, and supports click and
/// arrow/Home/End navigation. Disabling it cancels pending thumbnail renders.
/// Serve `public/page-thumbnails.js` beside the host entry, or set `runtime_url`.
/// Replacing or disposing the document automatically retires its navigator.
pub fn page_thumbnails_command(
  config : SurfaceConfig,
  target_id : String,
  enabled~ : Bool = true,
  runtime_url~ : String = "./page-thumbnails.js",
) -> @rabbita.Cmd {
  let kind = match config.kind {
    Some(Word) => "word"
    Some(Pdf) => "pdf"
    _ => return dispose_page_thumbnails_command(config.surface_id)
  }
  @cmd.custom_cmd(
    _ => {
      install_page_thumbnails(
        config.surface_id,
        target_id,
        kind,
        config.document_file,
        enabled,
        config.prefer_system_fonts,
        runtime_url,
      )
    },
    kind=@cmd.after_render,
  )
}

///|
/// Retire the navigator before removing its thumbnail container. Document
/// disposal also invokes this cleanup; repeated calls are harmless.
pub fn dispose_page_thumbnails_command(surface_id : String) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => retire_page_thumbnails(surface_id),
    kind=@cmd.after_render,
  )
}

///|
extern "js" fn install_page_thumbnails(
  surface_id : String,
  target_id : String,
  kind : String,
  key : String,
  enabled : Bool,
  system_fonts : Bool,
  runtime_url : String,
) =
  #| (surfaceId, targetId, kind, key, enabled, systemFonts, runtimeUrl) => {
  #|   const host = document.getElementById(surfaceId), target = document.getElementById(targetId);
  #|   if (!host || !target) return;
  #|   const request = host.__officeViewerPageThumbnailsRequest = (host.__officeViewerPageThumbnailsRequest || 0) + 1;
  #|   host.__officeViewerPageThumbnailsCleanup = () => {
  #|     host.__officeViewerPageThumbnailsRequest++;
  #|     host.__officeViewerPageThumbnails?.dispose();
  #|   };
  #|   import(new URL(runtimeUrl, document.baseURI).href).then(module => {
  #|     if (host.isConnected && target.isConnected && host.__officeViewerPageThumbnailsRequest === request) {
  #|       module.mount(host, target, {kind, key, enabled, systemFonts});
  #|     }
  #|   }).catch(() => {
  #|     if (host.__officeViewerPageThumbnailsRequest === request && target.isConnected) {
  #|       host.__officeViewerPageThumbnails?.dispose();
  #|       target.textContent = target.dataset.unavailableLabel;
  #|     }
  #|   });
  #| }

///|
extern "js" fn retire_page_thumbnails(surface_id : String) =
  #| id => document.getElementById(id)?.__officeViewerPageThumbnailsCleanup?.()
