/// Browser primitives only. Viewer policy, state, and UI construction live in
/// MoonBit; the bindings here expose APIs missing from Rabbita's DOM package.

///|
extern "js" fn node_text(node : @dom.Node) -> String = "node => node.textContent || ''"

///|
extern "js" fn node_set_text(node : @dom.Node, text : String) = "(node, text) => { node.textContent = text; }"

///|
extern "js" fn element_style(node : @dom.Element) -> @dom.CSSStyleDeclaration = "node => node.style"

///|
extern "js" fn computed_style(node : @dom.Element, name : String) -> String = "(node, name) => getComputedStyle(node).getPropertyValue(name)"

///|
extern "js" fn browser_number(value : String) -> Double = "value => Number(value)"

///|
extern "js" fn css_pixels(value : String) -> Double = "value => parseFloat(value) || 0"

///|
extern "js" fn pdf_current_zoom(host : @dom.Element) -> @js.Nullable[Double] =
  #| host => {
  #|   const session = host.__officeViewerPdf;
  #|   return session?.alive && session.doc && session.viewer.pagesCount ? session.viewer.currentScale * 100 : null;
  #| }

///|
extern "js" fn pdf_zoom_at(
  host : @dom.Element,
  zoom : Int32,
  x : Double,
  y : Double,
) =
  #| (host, zoom, x, y) => {
  #|   const viewer = host.__officeViewerPdf?.viewer;
  #|   if (!viewer?.pagesCount) return;
  #|   const bounds = viewer.container.getBoundingClientRect(), [top, left] = viewer.containerTopLeft;
  #|   viewer.updateScale({ scaleFactor: zoom / (viewer.currentScale * 100),
  #|     origin: [x - bounds.left + left, y - bounds.top + top], drawingDelay: 120 });
  #| }

///|
extern "js" fn get_boolean(node : @dom.Element, name : String) -> Bool = "(node, name) => !!node[name]"

///|
extern "js" fn set_boolean(node : @dom.Element, name : String, value : Bool) = "(node, name, value) => { node[name] = value; }"

///|
extern "js" fn get_callback(
  node : @dom.Element,
  name : String,
) -> @js.Nullable[() -> Unit] = "(node, name) => typeof node[name] === 'function' ? node[name] : null"

///|
extern "js" fn set_callback(
  node : @dom.Element,
  name : String,
  callback : () -> Unit,
) = "(node, name, callback) => { node[name] = callback; }"

///|
extern "js" fn annotation_hook(
  node : @dom.Element,
  name : String,
  callback : (String) -> Bool,
) =
  #| (node, name, callback) => {
  #|   node[name] = request => { try { return callback(JSON.stringify(request) ?? 'null'); } catch (_) { return false; } };
  #| }

///|
extern "js" fn fonts_ready(callback : () -> Unit) = "callback => { document.fonts?.ready.then(callback).catch(() => {}); }"

///|
extern "js" fn load_fonts(
  faces : FixedArray[String],
  callback : (Bool) -> Unit,
) =
  #| (faces, callback) => {
  #|   const loads = document.fonts ? faces.map(face => document.fonts.load(face)) : [];
  #|   Promise.all(loads).then(() => callback(true), () => callback(false));
  #| }

///|
extern "js" fn has_resize_observer() -> Bool = "() => typeof ResizeObserver === 'function'"

///|
extern "js" fn normalize_name(value : String) -> String = "value => value.trim().normalize('NFC')"

///|
extern "js" fn browser_graphemes(value : String) -> FixedArray[String] =
  #| value => typeof Intl.Segmenter === 'function'
  #|   ? Array.from(new Intl.Segmenter(undefined, {granularity: 'grapheme'}).segment(value), part => part.segment)
  #|   : Array.from(value)

///|
extern "js" fn locale_date(value : String) -> String =
  #| value => {
  #|   const date = new Date(value);
  #|   return Number.isNaN(date.getTime()) ? '' : new Intl.DateTimeFormat(undefined,
  #|     {year: 'numeric', month: 'short', day: 'numeric'}).format(date);
  #| }

///|
extern "js" fn uppercase(value : String) -> String = "value => value.toUpperCase()"

///|
extern "js" fn locale_lowercase(value : String) -> String = "value => value.toLocaleLowerCase()"

///|
#external
priv type NativeSelection

///|
#external
priv type NativeRange

///|
extern "js" fn native_selection() -> @js.Nullable[NativeSelection] = "() => window.getSelection()"

///|
extern "js" fn selection_count(selection : NativeSelection) -> Int32 = "s => s.rangeCount"

///|
extern "js" fn selection_collapsed(selection : NativeSelection) -> Bool = "s => s.isCollapsed"

///|
extern "js" fn selection_text(selection : NativeSelection) -> String = "s => s.toString()"

///|
extern "js" fn selection_anchor(
  selection : NativeSelection,
) -> @js.Nullable[@dom.Node] = "s => s.anchorNode"

///|
extern "js" fn selection_clear(selection : NativeSelection) = "s => s.removeAllRanges()"

///|
extern "js" fn selection_range(selection : NativeSelection) -> NativeRange = "s => s.getRangeAt(0)"

///|
extern "js" fn range_node(range : NativeRange, start : Bool) -> @dom.Node = "(r, start) => start ? r.startContainer : r.endContainer"

///|
extern "js" fn range_offset(range : NativeRange, start : Bool) -> Int32 = "(r, start) => start ? r.startOffset : r.endOffset"

///|
extern "js" fn range_intersects(range : NativeRange, node : @dom.Node) -> Bool = "(r, node) => r.intersectsNode(node)"

///|
extern "js" fn range_prefix_length(
  root : @dom.Node,
  node : @dom.Node,
  offset : Int32,
  fallback : Int32,
) -> Int32 =
  #| (root, node, offset, fallback) => {
  #|   const range = document.createRange();
  #|   range.selectNodeContents(root);
  #|   try { range.setEnd(node, offset); return range.toString().length; }
  #|   catch (_) { return fallback; }
  #| }

///|
extern "js" fn excel_reveal_hook(
  node : @dom.Element,
  reveal : (String, Int32, Int32) -> Bool,
) =
  #| (node, reveal) => { node.__officeRevealExcelCell = reveal; }

///|
extern "js" fn svg_screen_matrix(
  node : @dom.Element,
) -> @js.Nullable[@js.Value] = "node => typeof node.getScreenCTM === 'function' ? node.getScreenCTM() : null"

///|
extern "js" fn display_pixel_ratio() -> Double = "() => window.devicePixelRatio || 1"
