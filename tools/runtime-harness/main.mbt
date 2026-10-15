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
  expose_zoom_controls(
    @browser.configure_fit_width, @browser.install_pinch_zoom,
  )
}

///|
extern "js" fn expose_zoom_controls(
  fit : (String, String, Bool, (Int32) -> Unit) -> Unit,
  pinch : (String, String, Int32, Int32, (Int32) -> Unit) -> Unit,
) = "(fit, pinch) => { globalThis.officeTest.fit = fit; globalThis.officeTest.pinch = pinch; }"
