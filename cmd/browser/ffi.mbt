///|
// Browser-only helpers are kept behind a small boundary so the viewer model
// remains testable without a DOM.

///|
#cfg(target="js")
extern "js" fn selected_file_blob(input_id : String) -> @js.Nullable[@dom.Blob] =
  #| (id) => {
  #|   const input = document.getElementById(id);
  #|   return input && input.files && input.files.length > 0 ? input.files[0] : null;
  #| }

///|
#cfg(target="js")
extern "js" fn selected_file_name(input_id : String) -> String =
  #| (id) => {
  #|   const input = document.getElementById(id);
  #|   return input && input.files && input.files.length > 0 ? input.files[0].name : "";
  #| }

///|
#cfg(target="js")
extern "js" fn dropped_file_blob(
  event : @dom.DragEvent,
) -> @js.Nullable[@dom.Blob] =
  #| (event) => {
  #|   const files = event && event.dataTransfer && event.dataTransfer.files;
  #|   return files && files.length > 0 ? files[0] : null;
  #| }

///|
#cfg(target="js")
extern "js" fn dropped_file_name(event : @dom.DragEvent) -> String =
  #| (event) => {
  #|   const files = event && event.dataTransfer && event.dataTransfer.files;
  #|   return files && files.length > 0 ? files[0].name : "";
  #| }

///|
async fn read_selected_file(input_id : String) -> FilePayload? noraise {
  try {
    let blob = selected_file_blob(input_id).to_option()
    match blob {
      Some(blob) =>
        Some({ name: selected_file_name(input_id), bytes: blob.bytes(), })
      None => None
    }
  } catch {
    _ => None
  }
}

///|
async fn read_dropped_file(event : @dom.DragEvent) -> FilePayload? noraise {
  try {
    event.prevent_default()
    let blob = dropped_file_blob(event).to_option()
    match blob {
      Some(blob) =>
        Some({ name: dropped_file_name(event), bytes: blob.bytes(), })
      None => None
    }
  } catch {
    _ => None
  }
}

///|
#cfg(target="js")
extern "js" fn open_file_picker(input_id : String) =
  #| (id) => {
  #|   const input = document.getElementById(id);
  #|   if (input) input.click();
  #| }

///|
#cfg(target="js")
extern "js" fn set_surface_zoom(surface_id : String, zoom : Double) =
  #| (id, zoom) => {
  #|   const surface = document.getElementById(id);
  #|   if (surface) surface.style.setProperty("--office-zoom", String(zoom));
  #| }

///|
#cfg(target="js")
extern "js" fn set_system_font_mode(surface_id : String, enabled : Bool) =
  #| (id, enabled) => {
  #|   const surface = document.getElementById(id);
  #|   if (!surface) return;
  #|   surface.classList.toggle("prefer-system-fonts", enabled);
  #|   surface.dataset.officeFontMode = enabled ? "system" : "bundled";
  #| }

///|
#cfg(target="js")
extern "js" fn prepare_font_runtime(surface_id : String) =
  #| (id) => {
  #|   const surface = document.getElementById(id);
  #|   if (!surface) return;
  #|   const faces = [
  #|     ["normal", "400", "Carlito"],
  #|     ["normal", "700", "Carlito"],
  #|     ["italic", "400", "Carlito"],
  #|     ["italic", "700", "Carlito"],
  #|     ["normal", "400", "Liberation Sans"],
  #|     ["normal", "700", "Liberation Sans"],
  #|     ["italic", "400", "Liberation Sans"],
  #|     ["italic", "700", "Liberation Sans"],
  #|     ["normal", "400", "Liberation Serif"],
  #|     ["normal", "700", "Liberation Serif"],
  #|     ["italic", "400", "Liberation Serif"],
  #|     ["italic", "700", "Liberation Serif"],
  #|     ["normal", "400", "Liberation Mono"],
  #|     ["normal", "700", "Liberation Mono"],
  #|     ["italic", "400", "Liberation Mono"],
  #|     ["italic", "700", "Liberation Mono"],
  #|     ["normal", "400", "Noto Sans SC"],
  #|   ];
  #|   const fonts = document.fonts;
  #|   const ready = fonts
  #|     ? Promise.all(faces.map(([style, weight, family]) =>
  #|         fonts.load(`${style} ${weight} 12px "${family}"`)))
  #|     : Promise.resolve();
  #|   ready.then(() => {
  #|     surface.dataset.officeFonts = "bundled";
  #|   }).catch(() => {
  #|     surface.dataset.officeFonts = "fallback";
  #|   });
  #| }

///|
#cfg(target="js")
extern "js" fn apply_search(surface_id : String, query : String) =
  #| (id, query) => {
  #|   const surface = document.getElementById(id);
  #|   if (!surface) return;
  #|   const needle = query.trim().toLocaleLowerCase();
  #|   for (const node of surface.querySelectorAll(".office-search-hit")) {
  #|     node.classList.remove("office-search-hit");
  #|   }
  #|   if (!needle) return;
  #|   for (const node of surface.querySelectorAll("svg text, td")) {
  #|     if ((node.textContent || "").toLocaleLowerCase().includes(needle)) {
  #|       node.classList.add("office-search-hit");
  #|     }
  #|   }
  #| }

///|
#cfg(target="js")
extern "js" fn copy_to_clipboard(text : String) =
  #| (text) => {
  #|   if (text && navigator.clipboard && navigator.clipboard.writeText) {
  #|     navigator.clipboard.writeText(text).catch(() => {});
  #|   }
  #| }

///|
#cfg(target="js")
extern "js" fn install_surface_interactions(
  surface_id : String,
  kind : String,
  document_file : String,
  on_selection : (String) -> Unit,
) =
  #| (id, kind, documentFile, onSelection) => {
  #|   const host = document.getElementById(id);
  #|   if (!host) return;
  #|   if (host.__officeViewerCleanup) host.__officeViewerCleanup();
  #|   const cleanups = [];
  #|   const add = (target, type, handler, options) => {
  #|     target.addEventListener(type, handler, options);
  #|     cleanups.push(() => target.removeEventListener(type, handler, options));
  #|   };
  #|   const columnName = (number) => {
  #|     let value = number;
  #|     let name = "";
  #|     while (value > 0) {
  #|       const remainder = (value - 1) % 26;
  #|       name = String.fromCharCode(65 + remainder) + name;
  #|       value = Math.floor((value - 1) / 26);
  #|     }
  #|     return name;
  #|   };
  #|   const documentFormat = kind === "word" ? "docx" : "xlsx";
  #|   const jsonSelection = (value) => onSelection(JSON.stringify(value, null, 2));
  #|   const nodeIsInside = (root, node) => root === node || root.contains(node);
  #|
  #|   // SVG text is still selected through the browser's native Range. The
  #|   // source attributes on each text fragment let us turn that Range into
  #|   // document-local UTF-16 spans without changing the I-beam behavior.
  #|   const wordSelection = () => {
  #|     const selection = window.getSelection();
  #|     if (!selection || selection.rangeCount === 0) return;
  #|     const text = selection.toString();
  #|     if (!text.trim()) return;
  #|     const range = selection.getRangeAt(0);
  #|     if (!nodeIsInside(host, range.startContainer) ||
  #|         !nodeIsInside(host, range.endContainer)) return;
  #|     const anchors = [];
  #|     const byKey = new Map();
  #|     const pages = new Set();
  #|     const fragments = host.querySelectorAll("svg text[data-source-id]");
  #|     for (const fragment of fragments) {
  #|       const fragmentRange = document.createRange();
  #|       fragmentRange.selectNodeContents(fragment);
  #|       if (range.compareBoundaryPoints(Range.END_TO_START, fragmentRange) <= 0 ||
  #|           range.compareBoundaryPoints(Range.START_TO_END, fragmentRange) >= 0) {
  #|         continue;
  #|       }
  #|       const length = (fragment.textContent || "").length;
  #|       const offsetWithin = (container, offset, fallback) => {
  #|         if (!nodeIsInside(fragment, container)) return fallback;
  #|         const probe = document.createRange();
  #|         probe.selectNodeContents(fragment);
  #|         try {
  #|           probe.setEnd(container, offset);
  #|           return probe.toString().length;
  #|         } catch (_) {
  #|           return fallback;
  #|         }
  #|       };
  #|       const startsInside = nodeIsInside(fragment, range.startContainer);
  #|       const endsInside = nodeIsInside(fragment, range.endContainer);
  #|       const localStart = startsInside
  #|         ? offsetWithin(range.startContainer, range.startOffset, 0)
  #|         : 0;
  #|       const localEnd = endsInside
  #|         ? offsetWithin(range.endContainer, range.endOffset, length)
  #|         : length;
  #|       const sourceStart = Number(fragment.dataset.sourceStart || 0);
  #|       const sourceEnd = Number(fragment.dataset.sourceEnd || sourceStart + length);
  #|       const start = sourceStart + Math.max(0, Math.min(length, localStart));
  #|       const end = sourceStart + Math.max(0, Math.min(length, localEnd));
  #|       if (end <= start) continue;
  #|       const pageElement = fragment.closest(".word-page");
  #|       const page = pageElement && pageElement.dataset && pageElement.dataset.page
  #|         ? Number(pageElement.dataset.page)
  #|         : null;
  #|       if (page !== null) pages.add(page);
  #|       const id = fragment.dataset.sourceId || "";
  #|       const part = fragment.dataset.sourcePart || "";
  #|       const path = fragment.dataset.sourcePath || "";
  #|       const selector = fragment.dataset.sourceSelector || null;
  #|       const key = `${id}|${part}|${path}|${selector || ""}|${page === null ? "" : page}`;
  #|       const previous = byKey.get(key);
  #|       if (previous) {
  #|         previous.start_utf16 = Math.min(previous.start_utf16, start);
  #|         previous.end_utf16 = Math.max(previous.end_utf16, Math.min(sourceEnd, end));
  #|       } else {
  #|         const anchor = {
  #|           source_id: id,
  #|           part,
  #|           path,
  #|           selector,
  #|           stability: selector && selector.includes("[id=\"")
  #|             ? "stable"
  #|             : (selector ? "snapshot-relative" : "physical-only"),
  #|           start_utf16: start,
  #|           end_utf16: Math.min(sourceEnd, end),
  #|         };
  #|         if (page !== null) anchor.page = page;
  #|         byKey.set(key, anchor);
  #|         anchors.push(anchor);
  #|       }
  #|     }
  #|     anchors.sort((left, right) => {
  #|       const lp = left.page || 0;
  #|       const rp = right.page || 0;
  #|       return lp - rp || left.start_utf16 - right.start_utf16;
  #|     });
  #|     jsonSelection({
  #|       schema: "office.selection.v1",
  #|       file: documentFile,
  #|       format: documentFormat,
  #|       kind: "word",
  #|       text,
  #|       anchors,
  #|       pages: Array.from(pages).sort((left, right) => left - right),
  #|     });
  #|   };
  #|
  #|   const cellCoordinates = (cell) => {
  #|     if (!cell || !cell.dataset) return null;
  #|     const row = Number(cell.dataset.row);
  #|     const col = Number(cell.dataset.col);
  #|     if (!Number.isFinite(row) || !Number.isFinite(col)) return null;
  #|     return {
  #|       row,
  #|       col,
  #|       rowEnd: row + Math.max(1, cell.rowSpan || 1) - 1,
  #|       colEnd: col + Math.max(1, cell.colSpan || 1) - 1,
  #|     };
  #|   };
  #|   const cellFromPoint = (event) => {
  #|     const element = document.elementFromPoint(event.clientX, event.clientY);
  #|     return element && element.closest ? element.closest("td") : null;
  #|   };
  #|   const clearCells = (table) => {
  #|     if (!table) return;
  #|     for (const cell of table.querySelectorAll("td.office-cell-selected")) {
  #|       cell.classList.remove("office-cell-selected");
  #|     }
  #|   };
  #|   const selectCells = (anchor, active) => {
  #|     const table = anchor && anchor.closest("table");
  #|     if (!table || !active || active.closest("table") !== table) return null;
  #|     const a = cellCoordinates(anchor);
  #|     const b = cellCoordinates(active);
  #|     if (!a || !b) return null;
  #|     const top = Math.min(a.row, b.row);
  #|     const bottom = Math.max(a.rowEnd, b.rowEnd);
  #|     const left = Math.min(a.col, b.col);
  #|     const right = Math.max(a.colEnd, b.colEnd);
  #|     const selected = [];
  #|     clearCells(table);
  #|     for (const cell of table.querySelectorAll("td[data-cell]")) {
  #|       const coordinates = cellCoordinates(cell);
  #|       if (!coordinates) continue;
  #|       const intersects = coordinates.row <= bottom && coordinates.rowEnd >= top &&
  #|         coordinates.col <= right && coordinates.colEnd >= left;
  #|       if (intersects) {
  #|         cell.classList.add("office-cell-selected");
  #|         selected.push(cell);
  #|       }
  #|     }
  #|     selected.sort((leftCell, rightCell) => {
  #|       const leftCoordinates = cellCoordinates(leftCell);
  #|       const rightCoordinates = cellCoordinates(rightCell);
  #|       return leftCoordinates.row - rightCoordinates.row ||
  #|         leftCoordinates.col - rightCoordinates.col;
  #|     });
  #|     return { cells: selected, top, bottom, left, right };
  #|   };
  #|   const excelSelection = (pointer) => {
  #|     const selected = selectCells(pointer.anchor, pointer.active);
  #|     if (!selected || selected.cells.length === 0) return;
  #|     const table = pointer.table;
  #|     const section = table && table.closest ? table.closest("section") : null;
  #|     const sheet = section && section.dataset && section.dataset.sheet
  #|       ? section.dataset.sheet
  #|       : (selected.cells[0].dataset.sheet || "Sheet");
  #|     const rangeStart = `${columnName(selected.left)}${selected.top}`;
  #|     const rangeEnd = `${columnName(selected.right)}${selected.bottom}`;
  #|     const cells = selected.cells.map((cell) => {
  #|       const coordinates = cellCoordinates(cell);
  #|       return {
  #|         address: cell.dataset.cell,
  #|         row: coordinates.row,
  #|         column: coordinates.col,
  #|         text: cell.textContent || "",
  #|       };
  #|     });
  #|     jsonSelection({
  #|       schema: "office.selection.v1",
  #|       file: documentFile,
  #|       format: documentFormat,
  #|       kind: "excel",
  #|       sheet,
  #|       range: rangeStart === rangeEnd ? rangeStart : `${rangeStart}:${rangeEnd}`,
  #|       selector: `/xlsx/sheet[name=${JSON.stringify(sheet)}]/range[${rangeStart === rangeEnd ? rangeStart : `${rangeStart}:${rangeEnd}`}]`,
  #|       stability: "snapshot-relative",
  #|       cells,
  #|     });
  #|   };
  #|
  #|   let pointer = null;
  #|   const pointerDown = (event) => {
  #|     if (event.button !== undefined && event.button !== 0) return;
  #|     if (kind === "excel") {
  #|       const cell = event.target && event.target.closest ? event.target.closest("td") : null;
  #|       if (!cell || !cell.dataset.cell) return;
  #|       event.preventDefault();
  #|       pointer = { anchor: cell, active: cell, table: cell.closest("table") };
  #|       selectCells(cell, cell);
  #|       return;
  #|     }
  #|     if (!event.target || !event.target.closest || !event.target.closest(".word-page")) return;
  #|     pointer = { word: true };
  #|   };
  #|   const pointerMove = (event) => {
  #|     if (!pointer || kind !== "excel") return;
  #|     const cell = cellFromPoint(event);
  #|     if (cell) pointer.active = cell;
  #|     selectCells(pointer.anchor, pointer.active);
  #|   };
  #|   const pointerUp = () => {
  #|     if (!pointer) return;
  #|     if (kind === "excel") excelSelection(pointer);
  #|     else wordSelection();
  #|     pointer = null;
  #|   };
  #|   const pointerCancel = () => {
  #|     if (pointer && kind === "excel") clearCells(pointer.table);
  #|     pointer = null;
  #|   };
  #|   const selectionChanged = () => {
  #|     if (kind === "word") wordSelection();
  #|   };
  #|   add(host, "pointerdown", pointerDown);
  #|   add(window, "pointermove", pointerMove);
  #|   add(window, "pointerup", pointerUp);
  #|   add(window, "pointercancel", pointerCancel);
  #|   add(document, "selectionchange", selectionChanged);
  #|   if (kind === "excel") {
  #|     add(host, "dragstart", (event) => event.preventDefault());
  #|   }
  #|   host.__officeViewerCleanup = () => {
  #|     for (const cleanup of cleanups) cleanup();
  #|     host.__officeViewerCleanup = null;
  #|   };
  #| }
