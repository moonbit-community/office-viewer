///|
/// Stable PDF.js-owned subtree to pass to `SurfaceConfig.surface_html`.
pub fn pdf_surface_html() -> String {
  "<div class=\"pdf-container\" tabindex=\"0\"><div class=\"pdfViewer\"></div></div>"
}

///|
/// Load host-owned PDF bytes after `render_surface` has mounted a Pdf surface.
/// PDF.js, its worker, and supporting assets are served locally. State events
/// are JSON objects with a `status` (ready, page, scale, rotation, search,
/// password, loading, error, or selection-error). Selection events retain the
/// `office.selection.v1` envelope. The host must dispose before removing the
/// root; mounting another document automatically retires the previous runtime.
pub fn load_pdf_command(
  config : SurfaceConfig,
  bytes : Bytes,
  on_selection : (String) -> @rabbita.Cmd,
  on_state : (String) -> @rabbita.Cmd,
  runtime_url~ : String = "./pdf-runtime.js",
) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    scheduler => {
      mount_pdf(
        config.surface_id,
        config.document_file,
        bytes,
        config.zoom,
        config.query,
        runtime_url,
        value => scheduler.add(on_selection(value)),
        value => scheduler.add(on_state(value)),
      )
    },
    kind=@cmd.after_render,
  )
}

///|
/// Resolve a full, preview, deferred, or standalone PDF text-range reference
/// against the active surface. File path, SHA-256, extraction options, and the
/// PDF.js version must match. No cached selected text is needed. The result is
/// a full `office.selection.v1` envelope reconstructed from the PDF text items.
pub async fn resolve_pdf_selection(
  surface_id : String,
  selection_json : String,
) -> Result[String, String] noraise {
  Ok(pdf_resolve(surface_id, selection_json).wait()) catch {
    error => Err(error.to_string())
  }
}

///|
/// Read document identity, metadata, outline, and PDF permissions from the
/// active PDF.js document. Page text is extracted lazily as selections need it.
pub async fn pdf_document_info(
  surface_id : String,
) -> Result[String, String] noraise {
  Ok(pdf_info(surface_id).wait()) catch {
    error => Err(error.to_string())
  }
}

///|
pub fn pdf_page_command(surface_id : String, page : Int32) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => pdf_control(surface_id, "page", page.to_string()),
    kind=@cmd.after_render,
  )
}

///|
pub fn pdf_rotation_command(
  surface_id : String,
  rotation : Int32,
) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => pdf_control(surface_id, "rotation", rotation.to_string()),
    kind=@cmd.after_render,
  )
}

///|
pub fn pdf_fit_width_command(surface_id : String) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => pdf_control(surface_id, "fit", ""),
    kind=@cmd.after_render,
  )
}

///|
pub fn pdf_find_command(surface_id : String, previous : Bool) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => {
      pdf_control(
        surface_id,
        "find",
        if previous {
          "previous"
        } else {
          "next"
        },
      )
    },
    kind=@cmd.after_render,
  )
}

///|
pub fn pdf_password_command(
  surface_id : String,
  password : String,
) -> @rabbita.Cmd {
  @cmd.custom_cmd(
    _ => pdf_control(surface_id, "password", password),
    kind=@cmd.after_render,
  )
}

///|
#cfg(target="js")
extern "js" fn mount_pdf(
  surface_id : String,
  file : String,
  bytes : Bytes,
  zoom : Int32,
  query : String,
  runtime_url : String,
  on_selection : (String) -> Unit,
  on_state : (String) -> Unit,
) =
  #| (id, file, bytes, zoom, query, runtimeUrl, onSelection, onState) => {
  #|   const host = document.getElementById(id);
  #|   if (!host) return;
  #|   host.__officeViewerCleanup?.();
  #|   let cancelled = false, session = null;
  #|   const runtime = import(new URL(runtimeUrl, document.baseURI).href);
  #|   host.__officeViewerPdfRuntime = runtime;
  #|   host.__officeViewerCleanup = () => {
  #|     cancelled = true;
  #|     session?.destroy();
  #|     host.__officeViewerPdf = null;
  #|     host.__officeViewerPdfRuntime = null;
  #|     host.__officeViewerCleanup = null;
  #|   };
  #|   runtime.then(module => {
  #|     if (cancelled || !host.isConnected) return;
  #|     session = module.mount(host, {file, bytes, zoom, query, onSelection, onState});
  #|     host.__officeViewerPdf = session;
  #|   }).catch(error => {
  #|     if (!cancelled) onState(JSON.stringify({status: "error", message: error.message || String(error)}));
  #|   });
  #| }

///|
#cfg(target="js")
extern "js" fn apply_pdf_visual(
  surface_id : String,
  zoom : Int32,
  query : String,
) =
  #| (id, zoom, query) => document.getElementById(id)?.__officeViewerPdf?.apply(zoom, query)

///|
#cfg(target="js")
extern "js" fn pdf_control(
  surface_id : String,
  operation : String,
  value : String,
) =
  #| (id, operation, value) => {
  #|   const session = document.getElementById(id)?.__officeViewerPdf;
  #|   if (!session?.alive) return;
  #|   if (operation === "password") { session.passwordReply?.(value); session.passwordReply = null; return; }
  #|   if (!session.doc || !session.viewer.pagesCount) return;
  #|   if (operation === "page") session.viewer.currentPageNumber = Math.max(1, Math.min(session.doc.numPages, Number(value) || 1));
  #|   if (operation === "rotation") session.viewer.pagesRotation = ((Number(value) % 360) + 360) % 360;
  #|   if (operation === "fit") session.viewer.currentScaleValue = "page-width";
  #|   if (operation === "find") session.findAgain(value === "previous");
  #| }

///|
#cfg(target="js")
extern "js" fn pdf_resolve(
  surface_id : String,
  selection_json : String,
) -> @js.Promise[String] =
  #| async (id, selection) => {
  #|   const host = document.getElementById(id);
  #|   if (!host?.__officeViewerPdfRuntime) throw new Error("No active PDF runtime.");
  #|   return (await host.__officeViewerPdfRuntime).resolveSelection(id, selection);
  #| }

///|
#cfg(target="js")
extern "js" fn pdf_info(surface_id : String) -> @js.Promise[String] =
  #| async id => {
  #|   const host = document.getElementById(id);
  #|   if (!host?.__officeViewerPdfRuntime) throw new Error("No active PDF runtime.");
  #|   return (await host.__officeViewerPdfRuntime).documentInfo(id);
  #| }
