/// Browser tests call the same compiled MoonBit package as the production
/// surface. The harness only exposes entry points; it contains no viewer logic.

///|
extern "js" fn expose(
  install : (
    String,
    String,
    String,
    String,
    (String) -> Unit,
    (String, Int32, Int32, String) -> Unit,
  ) -> Unit,
  freeze : (String) -> Unit,
  dispose : (String) -> Unit,
  focus : (String, String) -> Bool,
  select : (String, String) -> Bool,
  clear : (String) -> Unit,
  zoom : (String, Double) -> Unit,
  search : (String, String) -> Unit,
  avatar : (String) -> @dom.Element,
) =
  #| (install, freeze, dispose, focus, select, clear, zoom, search, avatar) => {
  #|   globalThis.officeTest = {install, freeze, dispose, focus, select, clear, zoom, search,
  #|     avatar: item => avatar(JSON.stringify(item))};
  #| }

///|
fn main {
  expose(
    @browser.install_surface_interactions,
    @browser.install_excel_freeze,
    @browser.dispose_surface_interactions,
    @browser.focus_surface_annotation,
    @browser.select_surface_annotation,
    @browser.clear_surface_selection,
    @browser.set_surface_zoom,
    @browser.apply_search,
    value => {
      @browser.annotation_avatar(@json.parse(value) catch { _ => Json::null() })
    },
  )
  expose_font_runtime(@browser.prepare_font_runtime)
  expose_test_excel_renderer()
  expose_zoom_controls(
    @browser.configure_fit_width, @browser.install_pinch_zoom,
  )
  if @dom.document().get_element_by_id("thumbnail-test-controls").to_option()
    is Some(_) {
    @rabbita.elmish(
      model=0,
      update=(count, request : (String, String), _) => {
        let config : @viewer.SurfaceConfig = {
          surface_id: "document",
          kind: Some(Word),
          document_file: "/demo.docx",
          surface_html: "",
          zoom: 100,
          query: "",
          prefer_system_fonts: false,
        }
        (
          count + 1,
          @viewer.page_thumbnails_command(
            config,
            request.0,
            runtime_url=request.1,
          ),
        )
      },
      view=(_, emit) => {
        @html.div([
          @html.button(
            on_click=emit(("gallery", "./page-thumbnails.js")),
            "Mount thumbnail API",
          ),
          @html.button(
            on_click=emit(("replacement-list", "./missing-thumbnails.js")),
            "Fail thumbnail import",
          ),
          @html.button(
            on_click=emit(("replacement-list", "./delayed-missing.js")),
            "Delay thumbnail import",
          ),
          @viewer.render_page_thumbnails("replacement-list"),
        ])
      },
    ).mount("thumbnail-test-controls")
  }
}

///|
extern "js" fn expose_zoom_controls(
  fit : (String, String, Bool, (Int32) -> Unit) -> Unit,
  pinch : (String, String, Int32, Int32, (Int32) -> Unit) -> Unit,
) = "(fit, pinch) => { globalThis.officeTest.fit = fit; globalThis.officeTest.pinch = pinch; }"

///|
extern "js" fn expose_excel_renderer(render : (Bytes) -> String) =
  #| render => { globalThis.officeTest.renderExcel = base64 => render(Uint8Array.from(atob(base64), c => c.charCodeAt(0))); }

///|
fn expose_test_excel_renderer() -> Unit {
  expose_excel_renderer(bytes => {
    match @viewer.render_document("large.xlsx", "/docs/large.xlsx", bytes) {
      Ok(document) => document.surface_html
      Err(error) => "ERROR: " + error
    }
  })
}

///|
extern "js" fn expose_font_runtime(prepare : (String) -> Unit) = "prepare => { globalThis.officeTest.prepareFonts = prepare; }"
