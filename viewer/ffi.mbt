/// Browser-only helpers shared by standalone and embedded hosts.

///|
#cfg(target="js")
extern "js" fn set_surface_zoom(surface_id : String, zoom : Double) =
  #| (id, zoom) => {
  #|   const surface = document.getElementById(id);
  #|   if (surface) {
  #|     surface.style.setProperty("--office-zoom", String(zoom));
  #|     if (surface.__officeLayoutAnnotations) surface.__officeLayoutAnnotations();
  #|   }
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
  surface_html : String,
  on_selection : (String) -> Unit,
  on_annotation_selection : (String, Int, Int, String) -> Unit,
) =
  #| (id, kind, documentFile, surfaceHtml, onSelection, onAnnotationSelection) => {
  #|   const host = document.getElementById(id);
  #|   if (!host) return;
  #|   if (host.__officeViewerCleanup) host.__officeViewerCleanup();
  #|   // RawHtml can remain unchanged when the same file is opened again.
  #|   // Restore the canonical surface before deriving coordinates, rather
  #|   // than re-reading our own added headers, wrappers, and Note markers.
  #|   host.innerHTML = surfaceHtml;
  #|   const cleanups = [];
  #|   let alive = true;
  #|   let closeAnnotationContext = () => {};
  #|   let cancelCellPointer = () => {};
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
  #|   const jsonSelection = value => { closeAnnotationContext(); onSelection(JSON.stringify(value, null, 2)); };
  #|   const nodeIsInside = (root, node) => root === node || root.contains(node);
  #|   const documentFormat = kind === "word" ? "docx" : "xlsx";
  #|
  #|   // Position worksheet drawings over their anchor cells using the
  #|   // reader-provided offsets and dimensions. The overlay is transparent
  #|   // to pointer events so cell range selection remains uninterrupted.
  #|   const applyExcelObjectLayout = () => {
  #|     if (kind !== "excel") return;
  #|     const node = host.querySelector(".office-object-metadata");
  #|     let metadata = null;
  #|     if (node) {
  #|       try { metadata = JSON.parse(node.textContent || "{}"); }
  #|       catch (_) { metadata = null; }
  #|     }
  #|     if (!metadata || !metadata.sheets) return;
  #|     const sheetDocument = host.querySelector(".sheet-document");
  #|     const zoom = sheetDocument
  #|       ? (Number(getComputedStyle(sheetDocument).zoom) || 1) : 1;
  #|     for (const section of host.querySelectorAll("section[data-sheet]")) {
  #|       const sheet = section.dataset.sheet || "Sheet";
  #|       const objects = Array.isArray(metadata.sheets[sheet])
  #|         ? metadata.sheets[sheet] : [];
  #|       const table = section.querySelector("table");
  #|       const overlay = section.querySelector(":scope > .sheet-objects");
  #|       if (!table || !overlay) continue;
  #|       section.classList.add("office-anchored-objects");
  #|       overlay.classList.add("office-object-overlay");
  #|       const sectionRect = section.getBoundingClientRect();
  #|       const cells = Array.from(section.querySelectorAll("td[data-cell]"));
  #|       const figures = Array.from(overlay.querySelectorAll(
  #|         "figure.image, figure.chart, figure.chart-placeholder",
  #|       ));
  #|       let maxRight = table.offsetWidth || table.getBoundingClientRect().width;
  #|       let maxBottom = table.offsetHeight || table.getBoundingClientRect().height;
  #|       figures.forEach((figure, index) => {
  #|         const info = objects[index];
  #|         if (!info) return;
  #|         const anchor = String(info.anchor || figure.dataset.anchor || "");
  #|         const cell = cells.find((candidate) => candidate.dataset.cell === anchor);
  #|         if (!cell) return;
  #|         const cellRect = cell.getBoundingClientRect();
  #|         const offsetX = Number(info.offset_x) || 0;
  #|         const offsetY = Number(info.offset_y) || 0;
  #|         const width = Math.max(1, Number(info.width_px) || 1);
  #|         const height = Math.max(1, Number(info.height_px) || 1);
  #|         const left = (cellRect.left - sectionRect.left) / zoom + offsetX;
  #|         const top = (cellRect.top - sectionRect.top) / zoom + offsetY;
  #|         figure.style.position = "absolute";
  #|         figure.style.left = `${left}px`;
  #|         figure.style.top = `${top}px`;
  #|         figure.style.width = `${width}px`;
  #|         figure.style.height = `${height}px`;
  #|         figure.style.boxSizing = "border-box";
  #|         figure.style.margin = "0";
  #|         figure.style.pointerEvents = "none";
  #|         const image = figure.querySelector("img");
  #|         if (image) {
  #|           image.style.display = "block";
  #|           image.style.width = "100%";
  #|           image.style.height = "100%";
  #|           image.style.objectFit = "contain";
  #|         }
  #|         const svg = figure.querySelector("svg");
  #|         if (svg) {
  #|           svg.style.display = "block";
  #|           svg.setAttribute("width", String(width));
  #|           svg.setAttribute("height", String(height));
  #|           svg.style.width = "100%";
  #|           svg.style.height = "100%";
  #|         }
  #|         maxRight = Math.max(maxRight, left + width);
  #|         maxBottom = Math.max(maxBottom, top + height);
  #|       });
  #|       section.style.minWidth = `${Math.ceil(maxRight)}px`;
  #|       section.style.minHeight = `${Math.ceil(maxBottom)}px`;
  #|     }
  #|   };
  #|
  #|   const updateExcelFormulaBar = (cell, range) => {
  #|     if (kind !== "excel") return;
  #|     const nameBox = host.querySelector("[data-excel-name]");
  #|     const valueBox = host.querySelector("[data-excel-value]");
  #|     if (nameBox) nameBox.textContent = range || (cell && cell.dataset.cell) || "";
  #|     if (!valueBox) return;
  #|     if (!cell) {
  #|       valueBox.textContent = "Select a cell to inspect its value";
  #|       valueBox.classList.add("is-placeholder");
  #|       return;
  #|     }
  #|     const formula = cell.dataset.formula || "";
  #|     const value = formula ? `=${formula}` : (cell.textContent || "");
  #|     valueBox.textContent = value || " ";
  #|     valueBox.classList.remove("is-placeholder");
  #|   };
  #|
  #|   const setExcelSheet = (sheet) => {
  #|     if (kind !== "excel") return;
  #|     closeAnnotationContext();
  #|     cancelCellPointer();
  #|     const sections = host.querySelectorAll("section[data-sheet]");
  #|     let active = null;
  #|     for (const section of sections) {
  #|       const isActive = section.dataset.sheet === sheet;
  #|       section.classList.toggle("is-active", isActive);
  #|       section.hidden = !isActive;
  #|       if (isActive) active = section;
  #|     }
  #|     for (const tab of host.querySelectorAll("[data-excel-sheet-tab]")) {
  #|       const isActive = tab.dataset.excelSheetTab === sheet;
  #|       tab.classList.toggle("is-active", isActive);
  #|       tab.setAttribute("aria-selected", isActive ? "true" : "false");
  #|     }
  #|     host.dataset.officeActiveSheet = sheet || "";
  #|     for (const cell of host.querySelectorAll("td.office-cell-selected, td.office-cell-active")) {
  #|       cell.classList.remove("office-cell-selected", "office-cell-active");
  #|     }
  #|     updateExcelFormulaBar(null, "");
  #|     if (active) {
  #|       const table = active.querySelector("table");
  #|       if (table) table.scrollLeft = 0;
  #|     }
  #|     applyExcelObjectLayout();
  #|     const commentRail = host.querySelector(".office-context-rail");
  #|     if (commentRail) commentRail.scrollTop = 0;
  #|     if (host.__officeLayoutAnnotations) host.__officeLayoutAnnotations();
  #|   };
  #|
  #|   const enhanceExcelSurface = () => {
  #|     if (kind !== "excel") return;
  #|     const workbookSurface = host.querySelector(".sheet-document");
  #|     if (!workbookSurface) return;
  #|     const sections = Array.from(workbookSurface.querySelectorAll("section[data-sheet]"));
  #|     if (sections.length === 0) return;
  #|
  #|     const workspace = document.createElement("div");
  #|     workspace.className = "excel-workspace";
  #|     const toolbar = document.createElement("div");
  #|     toolbar.className = "excel-view-toolbar";
  #|     toolbar.innerHTML =
  #|       '<span class="excel-view-title">Spreadsheet</span>' +
  #|       '<span class="excel-readonly-badge">Read only</span>' +
  #|       '<span class="excel-view-hint">Select cells to give an Agent precise context</span>';
  #|     const formulaBar = document.createElement("div");
  #|     formulaBar.className = "excel-formula-bar";
  #|     formulaBar.innerHTML =
  #|       '<span class="excel-name-box" data-excel-name aria-label="Cell reference"></span>' +
  #|       '<span class="excel-fx" aria-hidden="true">fx</span>' +
  #|       '<span class="excel-value-box is-placeholder" data-excel-value>Select a cell to inspect its value</span>';
  #|     const tabs = document.createElement("nav");
  #|     tabs.className = "excel-sheet-tabs";
  #|     tabs.setAttribute("aria-label", "Worksheets");
  #|     const tabsLabel = document.createElement("span");
  #|     tabsLabel.className = "excel-tabs-label";
  #|     tabsLabel.textContent = "Sheets";
  #|     tabs.appendChild(tabsLabel);
  #|     const tabList = document.createElement("div");
  #|     tabList.className = "excel-tab-list";
  #|     tabs.appendChild(tabList);
  #|
  #|     workbookSurface.classList.add("office-excel-workbook");
  #|     const title = workbookSurface.querySelector(":scope > h1");
  #|     if (title) title.classList.add("excel-document-title");
  #|     for (const section of sections) {
  #|       section.classList.add("excel-sheet");
  #|       const heading = section.querySelector(":scope > h2");
  #|       if (heading) heading.classList.add("excel-sheet-heading");
  #|       const sheetName = section.dataset.sheet || "Sheet";
  #|       const tab = document.createElement("button");
  #|       tab.type = "button";
  #|       tab.className = "excel-sheet-tab";
  #|       tab.dataset.excelSheetTab = sheetName;
  #|       tab.setAttribute("role", "tab");
  #|       tab.setAttribute("aria-selected", "false");
  #|       tab.textContent = sheetName;
  #|       tabList.appendChild(tab);
  #|       add(tab, "click", (event) => {
  #|         event.preventDefault();
  #|         setExcelSheet(sheetName);
  #|       });
  #|     }
  #|
  #|     const parent = workbookSurface.parentNode;
  #|     if (!parent) return;
  #|     parent.insertBefore(workspace, workbookSurface);
  #|     workspace.appendChild(toolbar);
  #|     workspace.appendChild(formulaBar);
  #|     workspace.appendChild(workbookSurface);
  #|     workspace.appendChild(tabs);
  #|     const rememberedSheet = host.dataset.officeActiveSheet || "";
  #|     const hasRememberedSheet = sections.some(section => section.dataset.sheet === rememberedSheet);
  #|     setExcelSheet(hasRememberedSheet ? rememberedSheet : (sections[0].dataset.sheet || "Sheet"));
  #|   };
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
  #|   const applyExcelFormulaMetadata = () => {
  #|     if (kind !== "excel") return;
  #|     const node = host.querySelector(".office-formula-metadata");
  #|     if (!node) return;
  #|     let metadata = null;
  #|     try { metadata = JSON.parse(node.textContent || "{}"); }
  #|     catch (_) { metadata = null; }
  #|     if (!metadata || !metadata.sheets) return;
  #|     for (const section of host.querySelectorAll("section[data-sheet]")) {
  #|       const sheet = section.dataset.sheet || "Sheet";
  #|       const formulas = metadata.sheets[sheet] || {};
  #|       for (const cell of section.querySelectorAll("td[data-cell]")) {
  #|         const formula = formulas[cell.dataset.cell];
  #|         if (typeof formula === "string") cell.dataset.formula = formula;
  #|       }
  #|     }
  #|     node.remove();
  #|   };
  #|   const decorateExcelGrid = (table) => {
  #|     if (!table || table.dataset.officeGridDecorated === "true") return;
  #|     const colgroup = table.querySelector("colgroup");
  #|     const rows = Array.from(table.rows);
  #|     let columnCount = colgroup ? colgroup.querySelectorAll("col").length : 0;
  #|     if (columnCount === 0) {
  #|       for (const row of rows) {
  #|         let width = 0;
  #|         for (const cell of row.querySelectorAll("td"))
  #|           width += Math.max(1, cell.colSpan || 1);
  #|         columnCount = Math.max(columnCount, width);
  #|       }
  #|     }
  #|     if (colgroup) {
  #|       const rowHeaderCol = document.createElement("col");
  #|       rowHeaderCol.className = "office-row-header-column";
  #|       rowHeaderCol.style.width = "44px";
  #|       colgroup.insertBefore(rowHeaderCol, colgroup.firstChild);
  #|     }
  #|     const head = document.createElement("thead");
  #|     head.className = "office-grid-column-headings";
  #|     const headRow = document.createElement("tr");
  #|     const corner = document.createElement("th");
  #|     corner.className = "office-grid-corner";
  #|     corner.setAttribute("aria-label", "Select all cells");
  #|     headRow.appendChild(corner);
  #|     for (let column = 1; column <= columnCount; column++) {
  #|       const heading = document.createElement("th");
  #|       heading.className = "office-column-heading";
  #|       heading.scope = "col";
  #|       heading.textContent = columnName(column);
  #|       headRow.appendChild(heading);
  #|     }
  #|     head.appendChild(headRow);
  #|     const body = table.tBodies[0];
  #|     if (body) {
  #|       for (const row of Array.from(body.rows)) {
  #|         const firstCell = row.querySelector("td[data-row]");
  #|         const rowNumber = firstCell ? firstCell.dataset.row : String(row.rowIndex);
  #|         const heading = document.createElement("th");
  #|         heading.className = "office-row-heading";
  #|         heading.scope = "row";
  #|         heading.textContent = rowNumber;
  #|         row.insertBefore(heading, row.firstChild);
  #|       }
  #|     }
  #|     table.insertBefore(head, body || table.firstChild);
  #|     for (const cell of table.querySelectorAll("td[data-cell]")) {
  #|       cell.setAttribute("role", "gridcell");
  #|     }
  #|     table.dataset.officeGridDecorated = "true";
  #|   };
  #|   ensureExcelCellMetadata();
  #|   applyExcelFormulaMetadata();
  #|   if (kind === "excel") {
  #|     for (const table of host.querySelectorAll("section[data-sheet] table"))
  #|       decorateExcelGrid(table);
  #|   }
  #|   enhanceExcelSurface();
  #|   applyExcelObjectLayout();
  #|
  #|   // All selection requests resolve back to this installed snapshot. Caller
  #|   // supplied body/anchor fields are never trusted, even with a matching id.
  #|   let annotations = [];
  #|   try {
  #|     const data = host.querySelector('.office-annotations-metadata');
  #|     const envelope = data ? JSON.parse(data.textContent || '{}') : {};
  #|     if (envelope.schema === 'office.annotations.v1' && Array.isArray(envelope.items))
  #|       annotations = envelope.items.filter(item => item && item.document &&
  #|         item.document.file === documentFile && typeof item.document.sha256 === 'string');
  #|   } catch (_) {}
  #|   const currentAnnotation = (request) => {
  #|     const identity = request && request.document;
  #|     if (!alive || !host.isConnected || !identity || typeof identity.file !== 'string' ||
  #|         typeof identity.sha256 !== 'string') return null;
  #|     const matches = annotations.filter(candidate => candidate.id === request.id &&
  #|       candidate.document.file === identity.file && candidate.document.sha256 === identity.sha256);
  #|     return matches.length === 1 ? matches[0] : null;
  #|   };
  #|   const availableBody = item => item && typeof item.text === 'string' && item.text_status === 'available';
  #|   let selectedAnnotationId = null;
  #|   const emitAnnotation = (item, start, end, scope) => {
  #|     if (!currentAnnotation(item) || !availableBody(item) || !Number.isInteger(start) ||
  #|         !Number.isInteger(end) || start < 0 || end < start || end > item.text.length) return false;
  #|     const splitSurrogate = offset => offset > 0 && offset < item.text.length &&
  #|       item.text.charCodeAt(offset - 1) >= 0xD800 && item.text.charCodeAt(offset - 1) <= 0xDBFF &&
  #|       item.text.charCodeAt(offset) >= 0xDC00 && item.text.charCodeAt(offset) <= 0xDFFF;
  #|     if (splitSurrogate(start) || splitSurrogate(end)) return false;
  #|     selectedAnnotationId = item.id;
  #|     for (const entry of host.querySelectorAll('.office-context-entry'))
  #|       entry.classList.toggle('office-annotation-selected', entry.dataset.officeAnnotationId === item.id);
  #|     onAnnotationSelection(JSON.stringify(item), start, end, scope);
  #|     return true;
  #|   };
  #|   host.__officeSelectAnnotation = request => {
  #|     const item = currentAnnotation(request);
  #|     if (!item || !availableBody(item)) return false;
  #|     // A later selectionchange for an old native Range must not replace
  #|     // this explicit whole-body choice after the host renders its status.
  #|     window.getSelection()?.removeAllRanges();
  #|     return emitAnnotation(item, 0, item.text.length, 'whole');
  #|   };
  #|   const resolveAnnotationTargets = (item) => {
  #|     const anchor = item && item.anchor;
  #|     if (!anchor) return [];
  #|     if (kind === 'excel' && anchor.kind === 'xlsx') {
  #|       return Array.from(host.querySelectorAll('td[data-cell]')).filter(cell =>
  #|         cell.dataset.sheet === anchor.sheet && cell.dataset.cell === anchor.cell);
  #|     }
  #|     if (kind === 'word' && anchor.kind === 'docx' && Array.isArray(anchor.locations)) {
  #|       return Array.from(host.querySelectorAll('svg text[data-source-path]')).filter(node =>
  #|         anchor.locations.some(location => location.part === node.dataset.sourcePart &&
  #|           location.path === node.dataset.sourcePath));
  #|     }
  #|     return [];
  #|   };
  #|   const targetsByItem = new Map(annotations.map(item => [item, resolveAnnotationTargets(item)]));
  #|   const annotationTargets = item => targetsByItem.get(item) || [];
  #|   const annotationById = id => {
  #|     const matches = annotations.filter(item => item.id === id);
  #|     return matches.length === 1 ? matches[0] : null;
  #|   };
  #|   const threadRoot = item => {
  #|     const seen = new Set();
  #|     let current = item;
  #|     while (current && current.parent_id) {
  #|       if (seen.has(current.id)) return item;
  #|       seen.add(current.id);
  #|       const parent = annotationById(current.parent_id);
  #|       if (!parent || parent.document.sha256 !== item.document.sha256) break;
  #|       current = parent;
  #|     }
  #|     return current || item;
  #|   };
  #|   const contextFor = item => {
  #|     const own = annotationTargets(item);
  #|     if (own.length) return { item, targets: own, inherited: false };
  #|     const root = threadRoot(item);
  #|     const targets = root !== item ? annotationTargets(root) : [];
  #|     return { item: root, targets, inherited: targets.length > 0 };
  #|   };
  #|   const element = (tag, className, text) => {
  #|     const node = document.createElement(tag);
  #|     node.className = className;
  #|     if (text !== undefined) node.textContent = text;
  #|     return node;
  #|   };
  #|   // Comment bodies are present on first render. Focus is only a source-location
  #|   // aid; it never creates, hides or destroys the readable annotations.
  #|   const mapped = annotations.filter(item => annotationTargets(item).length > 0);
  #|   let rail = null, placed = null, unplaced = null, unplacedHeading = null;
  #|   let visibleCount = Math.min(200, annotations.length);
  #|   const cards = [], markerGroups = [];
  #|   const bodyItems = new WeakMap();
  #|   const clearContext = () => {
  #|     for (const node of host.querySelectorAll('.office-annotation-focus, .office-context-active'))
  #|       node.classList.remove('office-annotation-focus', 'office-context-active');
  #|   };
  #|   closeAnnotationContext = clearContext;
  #|   const visibleTarget = targets => targets.find(node => {
  #|     if (!node.isConnected || node.closest('section')?.hidden) return false;
  #|     if (kind !== 'excel') return true;
  #|     const viewport = node.closest('.office-excel-workbook');
  #|     if (!viewport) return true;
  #|     const r = node.getBoundingClientRect(), v = viewport.getBoundingClientRect();
  #|     // Non-layout hosts can still inspect the source mapping; actual browsers
  #|     // classify cells outside the scroll viewport as an explicit list fallback.
  #|     return !v.width || !v.height || (r.bottom > v.top && r.top < v.bottom && r.right > v.left && r.left < v.right);
  #|   });
  #|   const sourceLabel = item => {
  #|     const a = item.anchor || {};
  #|     if (a.kind === 'xlsx') return `${a.sheet || 'Unknown sheet'}!${a.cell || 'Unknown cell'}`;
  #|     const paths = (a.locations || []).map(location => location.path).filter(Boolean);
  #|     return paths.length ? `Paragraph ${paths.join(', ')} · context only` : 'Unplaced Word comment';
  #|   };
  #|   const positionAnnotations = () => {
  #|     if (!alive || !rail || !rail.isConnected) return;
  #|     const origin = placed.getBoundingClientRect().top + rail.scrollTop;
  #|     let bottom = 0, unplacedCount = 0;
  #|     const ordered = cards.map(record => ({ record, target: visibleTarget(record.context.targets) }))
  #|       .sort((a, b) => (a.target ? a.target.getBoundingClientRect().top : Infinity) -
  #|         (b.target ? b.target.getBoundingClientRect().top : Infinity));
  #|     for (const { record, target } of ordered) {
  #|       const { card, root, context, hint } = record;
  #|       const destination = target ? placed : unplaced;
  #|       if (card.parentNode !== destination) destination.appendChild(card);
  #|       card.classList.toggle('office-context-unplaced-card', !target);
  #|       if (target) {
  #|         const sourceTop = Math.max(0, target.getBoundingClientRect().top - origin);
  #|         const top = Math.max(sourceTop, bottom);
  #|         card.style.top = `${top}px`;
  #|         card.dataset.sourceTop = String(sourceTop);
  #|         card.style.setProperty('--office-card-offset', `${sourceTop - top}px`);
  #|         bottom = top + (card.offsetHeight || card.getBoundingClientRect().height || 1) + 12;
  #|         hint.textContent = `${sourceLabel(context.item)}${context.item !== root ? ` · Thread shown at ${context.item.id}'s source context.` : ''}`;
  #|       } else {
  #|         unplacedCount++;
  #|         card.style.top = '';
  #|         const a = context.item.anchor || {};
  #|         const reason = kind !== 'excel' ? 'No reliable rendered paragraph anchor.' :
  #|           a.sheet !== host.dataset.officeActiveSheet ? 'On another worksheet; body shown here now.' :
  #|           a.preview === 'outside-preview' ? 'Outside the rendered preview.' : 'Cell not in the current view.';
  #|         hint.textContent = `${sourceLabel(context.item)} · ${reason}`;
  #|       }
  #|     }
  #|     placed.style.height = `${bottom}px`;
  #|     unplacedHeading.hidden = unplacedCount === 0;
  #|     unplacedHeading.textContent = kind === 'word' ? `Unplaced comments (${unplacedCount} threads)` :
  #|       `Other locations (${unplacedCount} Notes) — bodies are shown below`;
  #|     let previousMarkerTop = -28;
  #|     for (const group of [...markerGroups].sort((a, b) => a.targets[0].getBoundingClientRect().top - b.targets[0].getBoundingClientRect().top)) {
  #|       const target = visibleTarget(group.targets);
  #|       group.button.hidden = !target;
  #|       if (target && kind === 'word') {
  #|         const sourceTop = Math.max(0, target.getBoundingClientRect().top - rail.getBoundingClientRect().top);
  #|         const top = Math.max(sourceTop, previousMarkerTop + 28);
  #|         group.button.style.top = `${top}px`;
  #|         group.button.style.setProperty('--office-marker-offset', `${sourceTop - top}px`);
  #|         previousMarkerTop = top;
  #|       }
  #|     }
  #|   };
  #|   let layoutFrame = null;
  #|   const scheduleAnnotationLayout = () => {
  #|     if (!alive || layoutFrame !== null) return;
  #|     layoutFrame = requestAnimationFrame(() => { layoutFrame = null; positionAnnotations(); });
  #|   };
  #|   host.__officeLayoutAnnotations = scheduleAnnotationLayout;
  #|   const showAnnotation = (item, _choices = [item], navigate = false) => {
  #|     if (!currentAnnotation(item) || !rail) return false;
  #|     const context = contextFor(item);
  #|     if (!context.targets.length) return false;
  #|     if (kind === 'excel' && host.dataset.officeActiveSheet !== context.item.anchor.sheet)
  #|       setExcelSheet(context.item.anchor.sheet);
  #|     clearContext();
  #|     for (const target of context.targets) target.classList.add('office-annotation-focus');
  #|     const record = cards.find(record => record.items.includes(item));
  #|     if (record) record.card.classList.add('office-context-active');
  #|     if (navigate) context.targets[0].scrollIntoView({ block: 'center', inline: 'center' });
  #|     scheduleAnnotationLayout();
  #|     return true;
  #|   };
  #|   const renderAnnotationCards = () => {
  #|     if (!rail) return;
  #|     for (const record of cards) record.card.remove();
  #|     cards.length = 0;
  #|     rail.querySelector('.office-context-more')?.remove();
  #|     const groups = new Map();
  #|     for (const item of annotations.slice(0, visibleCount)) {
  #|       const root = threadRoot(item);
  #|       if (!groups.has(root)) groups.set(root, []);
  #|       groups.get(root).push(item);
  #|     }
  #|     for (const [root, items] of groups) {
  #|       items.sort((a, b) => (a === root ? -1 : b === root ? 1 : 0));
  #|       // An anchored reply can supply visible thread context, explicitly labeled.
  #|       // Its source anchor is never copied onto the root or another reply.
  #|       const contextItem = annotationTargets(root).length ? root : items.find(item => annotationTargets(item).length) || root;
  #|       const context = { item: contextItem, targets: annotationTargets(contextItem) };
  #|       const card = element('aside', 'office-context-card');
  #|       card.dataset.officeThreadId = root.id;
  #|       card.setAttribute('aria-label', root.type === 'xlsx.note' ? 'Cell Note' : 'Comment thread');
  #|       const header = element('div', 'office-context-header');
  #|       header.appendChild(element('strong', '', root.type === 'xlsx.note' ? 'Note' : `Comment · ${root.id}`));
  #|       const locate = element('button', 'office-context-select', 'Show source');
  #|       locate.type = 'button';
  #|       locate.disabled = context.targets.length === 0;
  #|       locate.addEventListener('click', () => { if (card.isConnected) showAnnotation(contextItem, [contextItem], true); });
  #|       header.appendChild(locate);
  #|       card.appendChild(header);
  #|       const hint = element('div', 'office-context-location', sourceLabel(contextItem));
  #|       card.appendChild(hint);
  #|       for (const entry of items) {
  #|         const row = element('section', 'office-context-entry' + (entry !== root ? ' office-context-reply' : ''));
  #|         row.dataset.officeAnnotationId = entry.id;
  #|         row.classList.toggle('office-annotation-selected', entry.id === selectedAnnotationId);
  #|         row.appendChild(element('div', 'office-context-author', entry.author || 'Unknown author'));
  #|         row.appendChild(element('div', 'office-context-meta', [entry.date || 'Date unknown',
  #|           entry.resolved === true ? 'Resolved' : entry.resolved === false ? 'Open' : 'Resolution unknown'].join(' · ')));
  #|         if (entry !== root) row.appendChild(element('div', 'office-context-hint',
  #|           `Reply to ${entry.parent_id}. ${annotationTargets(entry).length ? 'Own source context recorded.' : 'No own rendered location; grouped with this thread.'}`));
  #|         const known = availableBody(entry);
  #|         const body = element('div', 'office-annotation-body', known ? entry.text : 'Comment text unavailable.');
  #|         body.dataset.officeAnnotationBody = entry.id;
  #|         body.setAttribute('tabindex', known ? '0' : '-1');
  #|         body.setAttribute('aria-label', `${entry !== root ? 'Reply' : 'Comment'} text by ${entry.author || 'Unknown author'}`);
  #|         if (known && !entry.text.length) body.dataset.empty = 'true';
  #|         if (known) bodyItems.set(body, entry);
  #|         body.addEventListener('click', () => {
  #|           if (!body.isConnected) return;
  #|           const selection = window.getSelection();
  #|           if (!selection || selection.isCollapsed) host.__officeSelectAnnotation(entry);
  #|         });
  #|         row.appendChild(body);
  #|         const select = element('button', 'office-context-select', entry.type === 'xlsx.note' ? 'Select Note' : 'Select comment');
  #|         select.type = 'button'; select.disabled = !known;
  #|         select.addEventListener('click', () => { if (select.isConnected) host.__officeSelectAnnotation(entry); });
  #|         row.appendChild(select);
  #|         card.appendChild(row);
  #|       }
  #|       placed.appendChild(card);
  #|       cards.push({ card, root, items, context, hint });
  #|     }
  #|     if (visibleCount < annotations.length) {
  #|       const more = element('div', 'office-context-more');
  #|       more.appendChild(element('p', '', `Showing ${visibleCount} of ${annotations.length} annotation bodies. Nothing is omitted from the structured data.`));
  #|       const button = element('button', 'office-context-select', `Show ${Math.min(200, annotations.length - visibleCount)} more`);
  #|       button.type = 'button';
  #|       button.addEventListener('click', () => { visibleCount = Math.min(annotations.length, visibleCount + 200); renderAnnotationCards(); });
  #|       more.appendChild(button); rail.appendChild(more);
  #|     }
  #|     positionAnnotations();
  #|     scheduleAnnotationLayout();
  #|   };
  #|   if (annotations.length) {
  #|     const content = host.querySelector(kind === 'word' ? '.word-pages' : '.excel-workspace') ||
  #|       (kind === 'word' ? host.querySelector('.word-page') : null);
  #|     if (content) {
  #|       const layout = element('div', `office-context-layout office-context-${kind}`);
  #|       const documentColumn = element('div', 'office-context-document');
  #|       rail = element('div', 'office-context-rail office-default-annotations');
  #|       rail.setAttribute('aria-label', 'Comments alongside the document');
  #|       host.insertBefore(layout, content);
  #|       layout.appendChild(documentColumn); documentColumn.appendChild(content); layout.appendChild(rail);
  #|       rail.appendChild(element('div', 'office-context-heading', `${annotations.length} ${kind === 'word' ? 'comments and replies' : 'Notes'} · select text directly below`));
  #|       placed = element('div', 'office-context-placed'); rail.appendChild(placed);
  #|       unplacedHeading = element('div', 'office-context-unplaced-heading'); rail.appendChild(unplacedHeading);
  #|       unplaced = element('div', 'office-context-unplaced'); rail.appendChild(unplaced);
  #|       cleanups.push(() => { if (layout.parentNode === host) host.insertBefore(content, layout); layout.remove(); });
  #|       renderAnnotationCards();
  #|     }
  #|   }
  #|   for (const item of mapped) {
  #|     for (const node of annotationTargets(item)) {
  #|       node.classList.add('office-has-annotation');
  #|       const ids = JSON.parse(node.dataset.officeAnnotationIds || '[]');
  #|       if (!ids.includes(item.id)) ids.push(item.id);
  #|       node.dataset.officeAnnotationIds = JSON.stringify(ids);
  #|       const page = node.closest('.word-page');
  #|       let group = markerGroups.find(group => kind === 'excel' ? group.targets[0] === node :
  #|         group.page === page && group.part === node.dataset.sourcePart && group.path === node.dataset.sourcePath);
  #|       if (!group) {
  #|         const button = element('button', kind === 'excel' ? 'office-note-marker' : 'office-comment-marker');
  #|         button.type = 'button';
  #|         group = { button, targets: [], items: [], page, part: node.dataset.sourcePart, path: node.dataset.sourcePath };
  #|         markerGroups.push(group);
  #|         if (kind === 'excel') node.appendChild(button); else if (rail) rail.appendChild(button);
  #|         const linkedGroup = group;
  #|         add(button, 'click', event => {
  #|           event.preventDefault(); event.stopPropagation();
  #|           const selected = linkedGroup.items[0];
  #|           if (showAnnotation(selected, linkedGroup.items)) host.__officeSelectAnnotation(selected);
  #|         });
  #|         cleanups.push(() => button.remove());
  #|       }
  #|       if (!group.targets.includes(node)) group.targets.push(node);
  #|       if (!group.items.includes(item)) group.items.push(item);
  #|       group.button.setAttribute('aria-label', `Select ${kind === 'excel' ? 'Note' : 'comment'} at ${kind === 'excel' ? node.dataset.cell : group.path}`);
  #|       group.button.setAttribute('title', 'Comment bodies are already shown alongside the document');
  #|     }
  #|   }
  #|   host.__officeFocusAnnotation = request => {
  #|     const item = currentAnnotation(request);
  #|     return !!item && showAnnotation(item, [item], true);
  #|   };
  #|   // Native Range offsets are measured against the exact plain text node(s),
  #|   // never Selection.toString(), which may normalize visible line endings.
  #|   const annotationSelection = () => {
  #|     if (!alive) return false;
  #|     const selection = window.getSelection();
  #|     if (!selection || selection.rangeCount !== 1 || selection.isCollapsed) return false;
  #|     const range = selection.getRangeAt(0);
  #|     const bodyFor = node => (node.nodeType === 1 ? node : node.parentElement)?.closest('[data-office-annotation-body]');
  #|     const startBody = bodyFor(range.startContainer), endBody = bodyFor(range.endContainer);
  #|     const intersectsCard = cards.some(record => range.intersectsNode(record.card));
  #|     if (!startBody || startBody !== endBody || !startBody.isConnected)
  #|       return !!(startBody || endBody || intersectsCard);
  #|     // Host-owned fallback bodies opt in with the same snapshot binding.
  #|     // This lets off-preview comments be selected without guessing anchors.
  #|     const item = host.contains(startBody) && bodyItems.has(startBody) ? bodyItems.get(startBody) :
  #|       startBody.dataset.officeAnnotationSurface === id ? currentAnnotation({
  #|         id: startBody.dataset.officeAnnotationBody,
  #|         document: { file: startBody.dataset.officeAnnotationFile, sha256: startBody.dataset.officeAnnotationSha256 },
  #|       }) : null;
  #|     if (!item || !availableBody(item)) return true;
  #|     const prefix = document.createRange();
  #|     prefix.selectNodeContents(startBody);
  #|     prefix.setEnd(range.startContainer, range.startOffset);
  #|     const start = prefix.toString().length;
  #|     prefix.setEnd(range.endContainer, range.endOffset);
  #|     const end = prefix.toString().length;
  #|     if (startBody.textContent !== item.text || end <= start) return true;
  #|     emitAnnotation(item, start, end, 'text');
  #|     return true;
  #|   };
  #|   add(window, 'keydown', event => { if (event.key === 'Escape') clearContext(); });
  #|   add(window, 'resize', scheduleAnnotationLayout);
  #|   add(document, 'scroll', event => { if (!rail || !rail.contains(event.target)) scheduleAnnotationLayout(); }, true);
  #|   if (typeof ResizeObserver !== 'undefined') {
  #|     const observer = new ResizeObserver(scheduleAnnotationLayout);
  #|     observer.observe(host);
  #|     if (rail) observer.observe(rail);
  #|     cleanups.push(() => observer.disconnect());
  #|   }
  #|   if (document.fonts && document.fonts.ready)
  #|     document.fonts.ready.then(() => { if (alive) { applyExcelObjectLayout(); scheduleAnnotationLayout(); } }).catch(() => {});
  #|   if (kind === 'excel') add(window, 'resize', applyExcelObjectLayout);
  #|   scheduleAnnotationLayout();
  #|   cleanups.push(() => {
  #|     if (layoutFrame !== null) cancelAnimationFrame(layoutFrame);
  #|     clearContext();
  #|     delete host.__officeFocusAnnotation;
  #|     delete host.__officeSelectAnnotation;
  #|     delete host.__officeLayoutAnnotations;
  #|   });
  #|
  #|   // SVG text is selected with the browser's native Range. Reader-backed
  #|   // metadata is trusted only when its paragraph-relative span covers the
  #|   // complete emitted fragment; otherwise the anchor stays physical-only.
  #|   const wordSelection = () => {
  #|     if (!alive || annotationSelection()) return;
  #|     const selection = window.getSelection();
  #|     if (!selection || selection.rangeCount !== 1) return;
  #|     const text = selection.toString();
  #|     if (!text.trim()) return;
  #|     const range = selection.getRangeAt(0);
  #|     if (!nodeIsInside(host, range.startContainer) ||
  #|         !nodeIsInside(host, range.endContainer)) return;
  #|     const pageFor = node => (node.nodeType === 1 ? node : node.parentElement)?.closest(".word-page");
  #|     if (!pageFor(range.startContainer) || !pageFor(range.endContainer)) return;
  #|     const anchors = [];
  #|     const byKey = new Map();
  #|     const pages = new Set();
  #|     const fragments = host.querySelectorAll("svg text");
  #|     for (const fragment of fragments) {
  #|       // `compareBoundaryPoints` is easy to reverse here because its
  #|       // argument is the source range. `intersectsNode` expresses the
  #|       // selection test directly and works for SVG text nodes as well.
  #|       if (!range.intersectsNode(fragment)) continue;
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
  #|     for (const cell of table.querySelectorAll("td.office-cell-selected, td.office-cell-active"))
  #|       cell.classList.remove("office-cell-selected", "office-cell-active");
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
  #|     active.classList.add("office-cell-active");
  #|     const rangeStart = `${columnName(left)}${top}`;
  #|     const rangeEnd = `${columnName(right)}${bottom}`;
  #|     updateExcelFormulaBar(active, rangeStart === rangeEnd ? rangeStart : `${rangeStart}:${rangeEnd}`);
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
  #|     if (!alive || (event.button !== undefined && event.button !== 0)) return;
  #|     if (event.target?.closest?.(".office-context-rail, .office-note-marker")) return;
  #|     if (kind === "excel") {
  #|       const cell = event.target && event.target.closest ? event.target.closest("td") : null;
  #|       if (!cell || !cell.dataset.cell) return;
  #|       clearContext();
  #|       event.preventDefault();
  #|       pointer = { anchor: cell, active: cell, table: cell.closest("table") };
  #|       selectCells(cell, cell);
  #|       return;
  #|     }
  #|     if (event.target && event.target.closest && event.target.closest(".word-page")) {
  #|       clearContext();
  #|       pointer = { word: true };
  #|     }
  #|   };
  #|   const pointerMove = (event) => {
  #|     if (!pointer || kind !== "excel") return;
  #|     const cell = cellFromPoint(event);
  #|     if (cell) pointer.active = cell;
  #|     selectCells(pointer.anchor, pointer.active);
  #|   };
  #|   const pointerUp = event => {
  #|     if (!alive || !pointer) return;
  #|     if (event.target?.closest?.('.office-context-rail, .office-note-marker')) { pointer = null; return; }
  #|     if (kind === 'excel') excelSelection(pointer); else wordSelection();
  #|     pointer = null;
  #|   };
  #|   const pointerCancel = () => { if (pointer && kind === "excel") clearCells(pointer.table); pointer = null; };
  #|   cancelCellPointer = pointerCancel;
  #|   const selectionChanged = () => { if (!annotationSelection() && kind === "word") wordSelection(); };
  #|   add(host, "pointerdown", pointerDown);
  #|   add(window, "pointermove", pointerMove);
  #|   add(window, "pointerup", pointerUp);
  #|   add(window, "pointercancel", pointerCancel);
  #|   add(document, "selectionchange", selectionChanged);
  #|   if (kind === "excel") add(host, "dragstart", (event) => event.preventDefault());
  #|   host.__officeViewerCleanup = () => { alive = false; pointer = null; for (const cleanup of cleanups) cleanup(); host.__officeViewerCleanup = null; };
  #| }

///|
#cfg(target="js")
extern "js" fn dispose_surface_interactions(surface_id : String) =
  #| (id) => {
  #|   const host = document.getElementById(id);
  #|   if (host && host.__officeViewerCleanup) host.__officeViewerCleanup();
  #| }

///|
#cfg(target="js")
extern "js" fn focus_surface_annotation(
  surface_id : String,
  item_json : String,
) -> Bool =
  #| (id, itemJson) => {
  #|   const host = document.getElementById(id);
  #|   if (!host || !host.__officeFocusAnnotation) return false;
  #|   try { return host.__officeFocusAnnotation(JSON.parse(itemJson)); }
  #|   catch (_) { return false; }
  #| }

///|
#cfg(target="js")
extern "js" fn select_surface_annotation(
  surface_id : String,
  item_json : String,
) -> Bool =
  #| (id, itemJson) => {
  #|   const host = document.getElementById(id);
  #|   if (!host || !host.__officeSelectAnnotation) return false;
  #|   try { return host.__officeSelectAnnotation(JSON.parse(itemJson)); }
  #|   catch (_) { return false; }
  #| }

///|
#cfg(target="js")
extern "js" fn clear_surface_selection(surface_id : String) =
  #| id => {
  #|   const host = document.getElementById(id);
  #|   const selection = window.getSelection();
  #|   if (host && selection?.anchorNode && host.contains(selection.anchorNode)) selection.removeAllRanges();
  #|   for (const cell of host?.querySelectorAll('.office-cell-selected, .office-cell-active') || []) {
  #|     cell.classList.remove('office-cell-selected', 'office-cell-active');
  #|   }
  #| }
