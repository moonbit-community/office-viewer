/// Browser-only helpers shared by standalone and embedded hosts.

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
  #|     ["normal", "400", "Carlito"], ["normal", "700", "Carlito"],
  #|     ["italic", "400", "Carlito"], ["italic", "700", "Carlito"],
  #|     ["normal", "400", "Liberation Sans"], ["normal", "700", "Liberation Sans"],
  #|     ["italic", "400", "Liberation Sans"], ["italic", "700", "Liberation Sans"],
  #|     ["normal", "400", "Liberation Serif"], ["normal", "700", "Liberation Serif"],
  #|     ["italic", "400", "Liberation Serif"], ["italic", "700", "Liberation Serif"],
  #|     ["normal", "400", "Liberation Mono"], ["normal", "700", "Liberation Mono"],
  #|     ["italic", "400", "Liberation Mono"], ["italic", "700", "Liberation Mono"],
  #|     ["normal", "400", "Noto Sans SC"],
  #|   ];
  #|   const fonts = document.fonts;
  #|   const ready = fonts
  #|     ? Promise.all(faces.map(([style, weight, family]) =>
  #|         fonts.load(`${style} ${weight} 12px "${family}"`)))
  #|     : Promise.resolve();
  #|   ready.then(() => { surface.dataset.officeFonts = "bundled"; })
  #|     .catch(() => { surface.dataset.officeFonts = "fallback"; });
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
  #|   const jsonSelection = (value) => onSelection(JSON.stringify(value, null, 2));
  #|   const nodeIsInside = (root, node) => root === node || root.contains(node);
  #|   const documentFormat = kind === "word" ? "docx" : "xlsx";
  #|
  #|   // The upstream xlsx2html renderer intentionally emits semantic sheet
  #|   // and table markup but does not know about browser selection. Build the
  #|   // address map here from actual DOM row/column coordinates. This keeps
  #|   // office.mbt untouched while preserving its merged-cell structure.
  #|   const ensureExcelCellMetadata = () => {
  #|     if (kind !== "excel") return;
  #|     for (const section of host.querySelectorAll("section[data-sheet]")) {
  #|       const sheet = section.dataset.sheet || "Sheet";
  #|       for (const table of section.querySelectorAll("table")) {
  #|         const occupied = new Map();
  #|         const rows = table.querySelectorAll("tr");
  #|         rows.forEach((rowElement, rowIndex) => {
  #|           let col = 1;
  #|           for (const cell of rowElement.querySelectorAll("td")) {
  #|             while (occupied.has(`${rowIndex}:${col}`)) col++;
  #|             const row = rowIndex + 1;
  #|             const colspan = Math.max(1, cell.colSpan || 1);
  #|             const rowspan = Math.max(1, cell.rowSpan || 1);
  #|             const address = `${columnName(col)}${row}`;
  #|             cell.dataset.sheet = sheet;
  #|             cell.dataset.row = String(row);
  #|             cell.dataset.col = String(col);
  #|             cell.dataset.cell = address;
  #|             for (let r = rowIndex; r < rowIndex + rowspan; r++) {
  #|               for (let c = col; c < col + colspan; c++) {
  #|                 occupied.set(`${r}:${c}`, true);
  #|               }
  #|             }
  #|             col += colspan;
  #|           }
  #|         });
  #|       }
  #|     }
  #|   };
  #|   ensureExcelCellMetadata();
  #|
  #|   // SVG text is selected with the browser's native Range. Reader-backed
  #|   // metadata is trusted only when its paragraph-relative span covers the
  #|   // complete emitted fragment; otherwise the anchor stays physical-only.
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
  #|     const fragments = host.querySelectorAll("svg text");
  #|     for (const fragment of fragments) {
  #|       const fragmentRange = document.createRange();
  #|       fragmentRange.selectNodeContents(fragment);
  #|       if (range.compareBoundaryPoints(Range.END_TO_START, fragmentRange) <= 0 ||
  #|           range.compareBoundaryPoints(Range.START_TO_END, fragmentRange) >= 0) continue;
  #|       const length = (fragment.textContent || "").length;
  #|       const offsetWithin = (container, offset, fallback) => {
  #|         if (!nodeIsInside(fragment, container)) return fallback;
  #|         const probe = document.createRange();
  #|         probe.selectNodeContents(fragment);
  #|         try { probe.setEnd(container, offset); return probe.toString().length; }
  #|         catch (_) { return fallback; }
  #|       };
  #|       const startsInside = nodeIsInside(fragment, range.startContainer);
  #|       const endsInside = nodeIsInside(fragment, range.endContainer);
  #|       const localStart = startsInside
  #|         ? offsetWithin(range.startContainer, range.startOffset, 0) : 0;
  #|       const localEnd = endsInside
  #|         ? offsetWithin(range.endContainer, range.endOffset, length) : length;
  #|       const start = Math.max(0, Math.min(length, localStart));
  #|       const end = Math.max(0, Math.min(length, localEnd));
  #|       if (end <= start) continue;
  #|       const pageElement = fragment.closest(".word-page");
  #|       const page = pageElement && pageElement.dataset && pageElement.dataset.page
  #|         ? Number(pageElement.dataset.page) : null;
  #|       if (page !== null) pages.add(page);
  #|       const sourceId = fragment.dataset.sourceId || "";
  #|       const sourcePart = fragment.dataset.sourcePart || "";
  #|       const sourcePath = fragment.dataset.sourcePath || "";
  #|       const sourceSelector = fragment.dataset.sourceSelector || null;
  #|       const sourceParaId = fragment.dataset.sourceParaId || null;
  #|       const sourceAnchorStatus = fragment.dataset.sourceAnchorStatus || null;
  #|       const sourceActionable = fragment.dataset.sourceActionable === "true";
  #|       const sourceReason = fragment.dataset.sourceReason || "";
  #|       const sourceStability = fragment.dataset.sourceStability || null;
  #|       const sourceStart = Number(fragment.dataset.sourceStart);
  #|       const sourceEnd = Number(fragment.dataset.sourceEnd);
  #|       const sourceRunsValue = fragment.dataset.sourceRuns || "";
  #|       const hasLogicalSpan = !!sourceSelector &&
  #|         Number.isFinite(sourceStart) && Number.isFinite(sourceEnd) &&
  #|         sourceEnd >= sourceStart && sourceEnd - sourceStart === length;
  #|       const logicalStart = hasLogicalSpan ? sourceStart + start : start;
  #|       const logicalEnd = hasLogicalSpan ? sourceStart + end : end;
  #|       const sourceRuns = hasLogicalSpan && sourceRunsValue
  #|         ? sourceRunsValue.split(",").filter(value => value.length > 0)
  #|         : undefined;
  #|       const rect = fragment.getBoundingClientRect();
  #|       const key = `${sourceId}|${sourcePart}|${sourcePath}|${sourceSelector || ""}|${page || ""}|${rect.left}|${rect.top}`;
  #|       const previous = byKey.get(key);
  #|       if (previous) {
  #|         previous.start_utf16 = Math.min(previous.start_utf16, logicalStart);
  #|         previous.end_utf16 = Math.max(previous.end_utf16, logicalEnd);
  #|       } else {
  #|         const anchor = {
  #|           source_id: sourceId || undefined,
  #|           part: sourcePart || undefined,
  #|           path: hasLogicalSpan ? sourcePath : undefined,
  #|           selector: hasLogicalSpan ? sourceSelector : undefined,
  #|           stability: hasLogicalSpan
  #|             ? (sourceStability || "snapshot-relative") : "physical-only",
  #|           start_utf16: logicalStart,
  #|           end_utf16: logicalEnd,
  #|           text: (fragment.textContent || "").slice(start, end),
  #|           rect: { left: rect.left, top: rect.top, width: rect.width, height: rect.height },
  #|         };
  #|         if (sourceRuns && sourceRuns.length > 0) anchor.runs = sourceRuns;
  #|         if (hasLogicalSpan) {
  #|           anchor.actionable = sourceActionable;
  #|           anchor.paragraph_anchor_status = sourceAnchorStatus || "missing";
  #|           if (sourceReason) anchor.reason = sourceReason;
  #|           if (sourceParaId) anchor.para_id = sourceParaId;
  #|         }
  #|         if (page !== null) anchor.page = page;
  #|         byKey.set(key, anchor);
  #|         anchors.push(anchor);
  #|       }
  #|     }
  #|     anchors.sort((left, right) => (left.page || 0) - (right.page || 0) ||
  #|       left.rect.top - right.rect.top || left.rect.left - right.rect.left);
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
  #|     return { row, col, rowEnd: row + Math.max(1, cell.rowSpan || 1) - 1,
  #|       colEnd: col + Math.max(1, cell.colSpan || 1) - 1 };
  #|   };
  #|   const cellFromPoint = (event) => {
  #|     const element = document.elementFromPoint(event.clientX, event.clientY);
  #|     return element && element.closest ? element.closest("td") : null;
  #|   };
  #|   const clearCells = (table) => {
  #|     if (!table) return;
  #|     for (const cell of table.querySelectorAll("td.office-cell-selected"))
  #|       cell.classList.remove("office-cell-selected");
  #|   };
  #|   const selectCells = (anchor, active) => {
  #|     const table = anchor && anchor.closest("table");
  #|     if (!table || !active || active.closest("table") !== table) return null;
  #|     const a = cellCoordinates(anchor), b = cellCoordinates(active);
  #|     if (!a || !b) return null;
  #|     const top = Math.min(a.row, b.row), bottom = Math.max(a.rowEnd, b.rowEnd);
  #|     const left = Math.min(a.col, b.col), right = Math.max(a.colEnd, b.colEnd);
  #|     const selected = [];
  #|     clearCells(table);
  #|     for (const cell of table.querySelectorAll("td[data-cell]")) {
  #|       const coordinates = cellCoordinates(cell);
  #|       if (!coordinates) continue;
  #|       if (coordinates.row <= bottom && coordinates.rowEnd >= top &&
  #|           coordinates.col <= right && coordinates.colEnd >= left) {
  #|         cell.classList.add("office-cell-selected");
  #|         selected.push(cell);
  #|       }
  #|     }
  #|     selected.sort((l, r) => { const a = cellCoordinates(l), b = cellCoordinates(r);
  #|       return a.row - b.row || a.col - b.col; });
  #|     return { cells: selected, top, bottom, left, right };
  #|   };
  #|   const excelSelection = (pointer) => {
  #|     const selected = selectCells(pointer.anchor, pointer.active);
  #|     if (!selected || selected.cells.length === 0) return;
  #|     const table = pointer.table;
  #|     const section = table && table.closest ? table.closest("section") : null;
  #|     const sheet = section && section.dataset && section.dataset.sheet
  #|       ? section.dataset.sheet : (selected.cells[0].dataset.sheet || "Sheet");
  #|     const start = `${columnName(selected.left)}${selected.top}`;
  #|     const end = `${columnName(selected.right)}${selected.bottom}`;
  #|     const range = start === end ? start : `${start}:${end}`;
  #|     const cells = selected.cells.map((cell) => { const c = cellCoordinates(cell);
  #|       return { address: cell.dataset.cell, row: c.row, column: c.col,
  #|         text: cell.textContent || "" }; });
  #|     const leaf = start === end ? `cell[${start}]` : `range[${range}]`;
  #|     jsonSelection({ schema: "office.selection.v1", file: documentFile,
  #|       format: documentFormat, kind: "excel", sheet, range,
  #|       selector: `/xlsx/sheet[name=${JSON.stringify(sheet)}]/${leaf}`,
  #|       stability: "snapshot-relative", cells });
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
  #|     if (event.target && event.target.closest && event.target.closest(".word-page"))
  #|       pointer = { word: true };
  #|   };
  #|   const pointerMove = (event) => {
  #|     if (!pointer || kind !== "excel") return;
  #|     const cell = cellFromPoint(event);
  #|     if (cell) pointer.active = cell;
  #|     selectCells(pointer.anchor, pointer.active);
  #|   };
  #|   const pointerUp = () => { if (!pointer) return; if (kind === "excel") excelSelection(pointer); else wordSelection(); pointer = null; };
  #|   const pointerCancel = () => { if (pointer && kind === "excel") clearCells(pointer.table); pointer = null; };
  #|   const selectionChanged = () => { if (kind === "word") wordSelection(); };
  #|   add(host, "pointerdown", pointerDown);
  #|   add(window, "pointermove", pointerMove);
  #|   add(window, "pointerup", pointerUp);
  #|   add(window, "pointercancel", pointerCancel);
  #|   add(document, "selectionchange", selectionChanged);
  #|   if (kind === "excel") add(host, "dragstart", (event) => event.preventDefault());
  #|   host.__officeViewerCleanup = () => { for (const cleanup of cleanups) cleanup(); host.__officeViewerCleanup = null; };
  #| }
