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
  on_selection : (String) -> Unit,
) =
  #| (id, kind, onSelection) => {
  #|   const host = document.getElementById(id);
  #|   if (!host) return;
  #|   if (host.__officeViewerCleanup) host.__officeViewerCleanup();
  #|   const cleanups = [];
  #|   const add = (target, type, handler, options) => {
  #|     target.addEventListener(type, handler, options);
  #|     cleanups.push(() => target.removeEventListener(type, handler, options));
  #|   };
  #|   const cellFromPoint = (event) => {
  #|     const element = document.elementFromPoint(event.clientX, event.clientY);
  #|     return element && element.closest ? element.closest("td") : null;
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
  #|   const buildCellLabels = (table) => {
  #|     if (table.__officeViewerCellLabels) return table.__officeViewerCellLabels;
  #|     const occupied = [];
  #|     const labels = new Map();
  #|     for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex += 1) {
  #|       if (!occupied[rowIndex]) occupied[rowIndex] = [];
  #|       let colIndex = 0;
  #|       for (const cell of Array.from(table.rows[rowIndex].cells)) {
  #|         while (occupied[rowIndex][colIndex]) colIndex += 1;
  #|         const colspan = Math.max(1, cell.colSpan || 1);
  #|         const rowspan = Math.max(1, cell.rowSpan || 1);
  #|         labels.set(cell, `${columnName(colIndex + 1)}${rowIndex + 1}`);
  #|         for (let row = rowIndex; row < rowIndex + rowspan; row += 1) {
  #|           if (!occupied[row]) occupied[row] = [];
  #|           for (let col = colIndex; col < colIndex + colspan; col += 1) {
  #|             occupied[row][col] = true;
  #|           }
  #|         }
  #|         colIndex += colspan;
  #|       }
  #|     }
  #|     table.__officeViewerCellLabels = labels;
  #|     return labels;
  #|   };
  #|   const cellLabel = (cell) => {
  #|     if (!cell) return "";
  #|     if (cell.dataset && cell.dataset.cell) return cell.dataset.cell;
  #|     const table = cell.closest("table");
  #|     const labels = table ? buildCellLabels(table) : null;
  #|     if (labels && labels.has(cell)) return labels.get(cell);
  #|     const row = cell.parentElement;
  #|     if (!row || !table) return "cell";
  #|     const rowIndex = Array.from(table.rows).indexOf(row) + 1;
  #|     const colIndex = Array.from(row.cells).indexOf(cell) + 1;
  #|     return `R${rowIndex}C${colIndex}`;
  #|   };
  #|   const clearCells = (table) => {
  #|     if (!table) return;
  #|     for (const cell of table.querySelectorAll("td.office-cell-selected")) {
  #|       cell.classList.remove("office-cell-selected");
  #|     }
  #|   };
  #|   const selectCells = (anchor, active) => {
  #|     const table = anchor && anchor.closest("table");
  #|     if (!table || !active || active.closest("table") !== table) return [];
  #|     const a = anchor.getBoundingClientRect();
  #|     const b = active.getBoundingClientRect();
  #|     const left = Math.min(a.left, b.left);
  #|     const right = Math.max(a.right, b.right);
  #|     const top = Math.min(a.top, b.top);
  #|     const bottom = Math.max(a.bottom, b.bottom);
  #|     const selected = [];
  #|     clearCells(table);
  #|     for (const cell of table.querySelectorAll("td")) {
  #|       const rect = cell.getBoundingClientRect();
  #|       const cx = (rect.left + rect.right) / 2;
  #|       const cy = (rect.top + rect.bottom) / 2;
  #|       if (cx >= left && cx <= right && cy >= top && cy <= bottom) {
  #|         cell.classList.add("office-cell-selected");
  #|         selected.push(cell);
  #|       }
  #|     }
  #|     return selected;
  #|   };
  #|   let pointer = null;
  #|   const pointerDown = (event) => {
  #|     if (event.button !== undefined && event.button !== 0) return;
  #|     if (kind === "excel") {
  #|       const cell = event.target && event.target.closest ? event.target.closest("td") : null;
  #|       if (!cell) return;
  #|       event.preventDefault();
  #|       pointer = { anchor: cell, active: cell, table: cell.closest("table") };
  #|       selectCells(cell, cell);
  #|       return;
  #|     }
  #|     if (!event.target || !event.target.closest || !event.target.closest(".word-page")) return;
  #|     const page = event.target.closest(".word-page");
  #|     pointer = {
  #|       page,
  #|     };
  #|   };
  #|   const pointerMove = (event) => {
  #|     if (!pointer) return;
  #|     if (kind !== "excel") return;
  #|     const cell = cellFromPoint(event);
  #|     if (cell) pointer.active = cell;
  #|     selectCells(pointer.anchor, pointer.active);
  #|   };
  #|   const pointerUp = (event) => {
  #|     if (!pointer) return;
  #|     if (kind === "excel") {
  #|       const cells = selectCells(pointer.anchor, pointer.active);
  #|       const values = cells.map((cell) => cell.textContent || "");
  #|       const labels = cells.map(cellLabel);
  #|       const first = labels.length ? labels[0] : "cell";
  #|       const table = pointer.table;
  #|       const section = table && table.closest ? table.closest("section") : null;
  #|       const sheet = section && section.dataset && section.dataset.sheet
  #|         ? section.dataset.sheet
  #|         : "Sheet";
  #|       onSelection(`Excel ${sheet}!${first} (${cells.length} cells)\n${values.join("\t")}`);
  #|     } else {
  #|       const selection = window.getSelection();
  #|       const text = selection ? selection.toString().trim() : "";
  #|       let belongsToPage = true;
  #|       if (text && selection && selection.rangeCount > 0 && pointer.page) {
  #|         const range = selection.getRangeAt(0);
  #|         belongsToPage = pointer.page.contains(range.startContainer) ||
  #|           pointer.page.contains(range.endContainer);
  #|       }
  #|       if (text && belongsToPage) {
  #|         const pageNumber = pointer.page && pointer.page.dataset
  #|           ? pointer.page.dataset.page
  #|           : "?";
  #|         onSelection(`Word page ${pageNumber}\n${text}`);
  #|       }
  #|     }
  #|     pointer = null;
  #|   };
  #|   add(host, "pointerdown", pointerDown);
  #|   add(window, "pointermove", pointerMove);
  #|   add(window, "pointerup", pointerUp);
  #|   add(window, "pointercancel", pointerUp);
  #|   if (kind === "excel") {
  #|     add(host, "dragstart", (event) => event.preventDefault());
  #|   }
  #|   host.__officeViewerCleanup = () => {
  #|     for (const cleanup of cleanups) cleanup();
  #|     host.__officeViewerCleanup = null;
  #|   };
  #| }
