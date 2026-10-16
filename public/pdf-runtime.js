import * as pdfjs from "./vendor/pdfjs/build/pdf.js";

// The upstream viewer consumes pdfjsLib; the core must be loaded first.
const { PDFViewer, EventBus, PDFLinkService, PDFFindController, LinkTarget } =
  await import("./vendor/pdfjs/web/pdf_viewer.js");
pdfjs.GlobalWorkerOptions.workerSrc = new URL("./vendor/pdfjs/build/pdf.worker.js", import.meta.url).href;
const asset = path => new URL(`./vendor/pdfjs/${path}/`, import.meta.url).href;
const sessions = new Map();
const textParams = Object.freeze({ includeMarkedContent: true, disableNormalization: true });
const referenceSchema = "office.pdf.text-range.v1";
const joinSchema = "pdfjs-items-eol-v1";

function endpointOrder(a, b) {
  return a.page - b.page || a.item_index - b.item_index || a.offset_utf16 - b.offset_utf16;
}

function textSpan(div, range) {
  if (!range.intersectsNode(div)) return null;
  const length = div.textContent.length;
  const offset = (node, value, fallback) => {
    if (div !== node && !div.contains(node)) return fallback;
    const probe = document.createRange();
    probe.selectNodeContents(div);
    probe.setEnd(node, value);
    return probe.toString().length;
  };
  const start = Math.max(0, Math.min(length, offset(range.startContainer, range.startOffset, 0)));
  const end = Math.max(0, Math.min(length, offset(range.endContainer, range.endOffset, length)));
  return end > start ? { start, end } : null;
}

function subrange(div, start, end) {
  const walker = document.createTreeWalker(div, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  let count = 0, started = false, node;
  while ((node = walker.nextNode())) {
    const next = count + node.length;
    if (!started && start <= next) {
      range.setStart(node, start - count);
      started = true;
    }
    if (started && end <= next) {
      range.setEnd(node, end - count);
      return range;
    }
    count = next;
  }
  return null;
}

function pdfQuads(pageView, range) {
  if (!range || !pageView) return [];
  const canvas = pageView.div.querySelector(".canvasWrapper");
  if (!canvas) return [];
  const bounds = canvas.getBoundingClientRect();
  const { viewport } = pageView;
  const sx = viewport.width / bounds.width, sy = viewport.height / bounds.height;
  return [...range.getClientRects()].filter(r => r.width > 0 && r.height > 0).map(r =>
    [[r.left, r.top], [r.right, r.top], [r.right, r.bottom], [r.left, r.bottom]].map(([x, y]) =>
      viewport.convertToPdfPoint((x - bounds.left) * sx, (y - bounds.top) * sy)));
}

function itemQuad(item, style) {
  const [a, b, c, d, x, y] = item.transform;
  const ascent = style?.ascent ?? (style?.descent != null ? 1 + style.descent : 0.8);
  const descent = style?.descent ?? ascent - 1;
  const length = Math.hypot(a, b) || 1;
  if (style?.vertical) {
    const height = Math.hypot(c, d), dx = b / length * item.height, dy = -a / length * item.height;
    const nx = a / length * height, ny = b / length * height;
    return [[x + nx * ascent, y + ny * ascent], [x + dx + nx * ascent, y + dy + ny * ascent],
      [x + dx + nx * descent, y + dy + ny * descent], [x + nx * descent, y + ny * descent]];
  }
  const dx = a / length * item.width, dy = b / length * item.width;
  return [[x + c * ascent, y + d * ascent], [x + dx + c * ascent, y + dy + d * ascent],
    [x + dx + c * descent, y + dy + d * descent], [x + c * descent, y + d * descent]];
}

function boxOf(quads) {
  const points = quads.flat();
  return [Math.min(...points.map(p => p[0])), Math.min(...points.map(p => p[1])),
    Math.max(...points.map(p => p[0])), Math.max(...points.map(p => p[1]))];
}

function structurePaths(tree) {
  const paths = new Map();
  const walk = (node, roles, properties) => {
    const next = node.role ? [...roles, node.role] : roles;
    const props = { ...properties };
    if (node.alt) props.alt = node.alt;
    if (node.lang) props.lang = node.lang;
    if (node.type === "content" && node.id) paths.set(node.id, { role_path: next, ...props });
    for (const child of node.children || []) walk(child, next, props);
  };
  if (tree) walk(tree, [], {});
  return paths;
}

class PdfSession {
  constructor(host, options) {
    this.host = host;
    this.options = options;
    this.documentFile = options.file;
    this.alive = true;
    this.pages = new Map();
    this.abort = new AbortController();
    this.selectionRevision = 0;
    this.selectionPointer = null;
    this.selectionCancelled = false;
    this.bus = new EventBus();
    this.links = new PDFLinkService({ eventBus: this.bus,
      externalLinkTarget: LinkTarget.BLANK, externalLinkRel: "noopener noreferrer nofollow" });
    this.find = new PDFFindController({ eventBus: this.bus, linkService: this.links });
    const container = host.querySelector(".pdf-container");
    this.viewer = new PDFViewer({ container, viewer: container.firstElementChild,
      eventBus: this.bus, linkService: this.links, findController: this.find,
      textLayerMode: 1,
      annotationMode: pdfjs.AnnotationMode.ENABLE,
      annotationEditorMode: pdfjs.AnnotationEditorType.DISABLE,
      enableScripting: false, enableAutoLinking: true,
      enableDetailCanvas: false, maxCanvasPixels: 16777216,
      imageResourcesPath: asset("web/images"), abortSignal: this.abort.signal });
    this.links.setViewer(this.viewer);
    this.bus.on("pagechanging", event => this.state({ status: "page", page: event.pageNumber }));
    this.bus.on("scalechanging", event => this.state({ status: "scale", zoom: Math.round(event.scale * 100) }));
    this.bus.on("rotationchanging", event => this.state({ status: "rotation", rotation: event.pagesRotation }));
    this.bus.on("updatefindmatchescount", event => this.state({ status: "search", ...event.matchesCount }));
    this.bus.on("updatefindcontrolstate", event => this.state({ status: "search", state: event.state,
      current: event.matchesCount?.current || 0, total: event.matchesCount?.total || 0 }));
    this.bus.on("pagerendered", event => {
      if (event.error) this.state({ status: "error", message: event.error.message || String(event.error) });
      else host.dataset.pdfPainted = "true";
    });
    this.bus.on("textlayerrendered", event => {
      if (!event.error) this.mapping(event.pageNumber).catch(() => {});
    });
    const schedule = () => {
      clearTimeout(this.selectionTimer);
      this.selectionRevision++;
      if (!this.alive || this.selectionCancelled || this.selectionPointer !== null) return;
      this.selectionTimer = setTimeout(() => this.capture().catch(error => {
        if (this.alive) this.state({ status: "selection-error", message: error.message });
      }), 90);
    };
    document.addEventListener("selectionchange", schedule, { signal: this.abort.signal });
    host.addEventListener("pointerdown", event => {
      if (event.button !== 0 || event.isPrimary === false) return;
      this.selectionCancelled = false;
      this.selectionPointer = event.pointerId;
      clearTimeout(this.selectionTimer);
      this.selectionRevision++;
    }, { signal: this.abort.signal });
    // Release outside the surface must settle the drag as well.
    document.addEventListener("pointerup", event => {
      if (this.selectionPointer === null || this.selectionPointer !== event.pointerId) return;
      this.selectionPointer = null;
      schedule();
    }, { signal: this.abort.signal });
    document.addEventListener("pointercancel", event => {
      if (this.selectionPointer === null || this.selectionPointer !== event.pointerId) return;
      this.selectionPointer = null;
      this.selectionCancelled = true;
      clearTimeout(this.selectionTimer);
      this.selectionRevision++;
    }, { signal: this.abort.signal });
    host.addEventListener("keydown", () => {
      this.selectionCancelled = false;
    }, { signal: this.abort.signal });
    host.addEventListener("keyup", schedule, { signal: this.abort.signal });
    // PDFViewer observes container height, but its embedding host must refit
    // scale presets and update visible pages when panel animations resize it.
    this.resizeObserver = new ResizeObserver(() => {
      if (!this.alive || this.resizeFrame != null) return;
      this.resizeFrame = requestAnimationFrame(() => {
        this.resizeFrame = null;
        if (!this.alive || !this.doc || !this.viewer.pagesCount || container.clientWidth <= 40) return;
        const scale = this.viewer.currentScaleValue;
        if (scale && Number.isNaN(Number(scale))) this.viewer.currentScaleValue = scale;
        this.viewer.update();
      });
    });
    this.resizeObserver.observe(container);
  }

  state(value) {
    if (this.alive) this.options.onState(JSON.stringify(value));
  }

  async open(bytes) {
    const copy = Uint8Array.from(bytes);
    const digest = await crypto.subtle.digest("SHA-256", copy);
    if (!this.alive) return;
    this.identity = { sha256: [...new Uint8Array(digest)].map(x => x.toString(16).padStart(2, "0")).join(""),
      byte_length: copy.byteLength, parser: { name: "pdfjs", version: pdfjs.version } };
    this.task = pdfjs.getDocument({ data: copy, cMapUrl: asset("cmaps"), cMapPacked: true,
      standardFontDataUrl: asset("standard_fonts"), wasmUrl: asset("wasm"), iccUrl: asset("iccs"),
      isEvalSupported: false, enableXfa: false });
    this.task.onPassword = (reply, reason) => {
      this.passwordReply = reply;
      this.state({ status: "password", incorrect: reason === pdfjs.PasswordResponses.INCORRECT_PASSWORD });
    };
    this.task.onProgress = ({ loaded, total }) => this.state({ status: "loading", loaded, total });
    this.doc = await this.task.promise;
    if (!this.alive) return;
    this.identity.fingerprints = this.doc.fingerprints;
    this.identity.page_count = this.doc.numPages;
    const initialized = new Promise(resolve => this.bus.on("pagesinit", resolve, { once: true }));
    this.links.setDocument(this.doc);
    this.viewer.setDocument(this.doc);
    await initialized;
    if (!this.alive) return;
    this.viewer.currentScaleValue = this.viewer.container.clientWidth < 640 ? "page-width" : String(this.options.zoom / 100);
    this.viewer.update();
    this.apply(this.options.zoom, this.options.query, false);
    this.host.dataset.pdfReady = "true";
    this.state({ status: "ready", pages: this.doc.numPages, document: this.identity });
  }

  async pageData(number) {
    if (!this.alive || !this.doc) throw new Error("No active PDF document.");
    if (!this.pages.has(number)) this.pages.set(number, (async () => {
      const page = await this.doc.getPage(number);
      const [content, tree] = await Promise.all([page.getTextContent(textParams), page.getStructTree().catch(() => null)]);
      const semantics = structurePaths(tree), marked = [], entries = [];
      content.items.forEach((item, index) => {
        if (typeof item.str === "string") entries.push({ item, index,
          markedId: [...marked].reverse().find(Boolean) || null });
        else if (item.type === "beginMarkedContentProps" || item.type === "beginMarkedContent") marked.push(item.id || null);
        else if (item.type === "endMarkedContent") marked.pop();
      });
      return { page, content, entries, semantics, tagged: !!tree };
    })());
    return this.pages.get(number);
  }

  async mapping(number) {
    const data = await this.pageData(number);
    if (!this.alive) return null;
    const view = this.viewer.getPageView(number - 1), highlighter = view?.textLayer?.highlighter;
    const divs = highlighter?.textDivs, strings = highlighter?.textContentItemsStr;
    if (!divs || divs.length !== data.entries.length || strings.some((s, i) => s !== data.entries[i].item.str)) return null;
    const result = data.entries.map((entry, i) => ({ ...entry, div: divs[i] }));
    for (const entry of result) {
      entry.div.dataset.pdfItem = String(entry.index);
      entry.div.dataset.pdfPage = String(number);
    }
    return { view, data, entries: result };
  }

  async capture() {
    if (this.selectionCancelled || this.selectionPointer !== null) return;
    const selection = window.getSelection();
    if (!this.doc || !selection?.rangeCount || selection.isCollapsed || !selection.toString().trim()) return;
    const range = selection.getRangeAt(0).cloneRange();
    if (!this.host.contains(range.startContainer) || !this.host.contains(range.endContainer)) return;
    const revision = this.selectionRevision, nativeText = selection.toString(), selected = [];
    for (const element of this.host.querySelectorAll(".page")) {
      if (!range.intersectsNode(element)) continue;
      const number = Number(element.dataset.pageNumber);
      const mapping = await this.mapping(number);
      if (!mapping) continue;
      for (const entry of mapping.entries) {
        if (!entry.div.isConnected || entry.div.textContent !== entry.item.str) continue;
        const span = textSpan(entry.div, range);
        if (span) selected.push({ page: number, item_index: entry.index, ...span });
      }
    }
    if (!this.alive || revision !== this.selectionRevision || !selected.length) return;
    const first = selected[0], last = selected.at(-1);
    const reference = { schema: referenceSchema, file: this.documentFile, document: this.identity,
      text_content: { include_marked_content: true, disable_normalization: true }, join: joinSchema,
      start: { page: first.page, item_index: first.item_index, offset_utf16: first.start },
      end: { page: last.page, item_index: last.item_index, offset_utf16: last.end } };
    const result = await this.read(reference, range);
    if (!this.alive || revision !== this.selectionRevision) return;
    reference.text_utf16_length = result.text.length;
    this.options.onSelection(JSON.stringify({ ...result, browser_text: nativeText }, null, 2));
  }

  validate(reference) {
    if (!this.alive || !this.doc) throw new Error("No active PDF document.");
    if (reference?.schema !== referenceSchema || reference.join !== joinSchema) throw new Error("Unsupported PDF reference schema.");
    if (!reference.file || reference.file !== this.documentFile || reference.document?.sha256 !== this.identity.sha256 ||
        reference.document?.byte_length !== this.identity.byte_length || reference.document?.page_count !== this.doc.numPages ||
        reference.document?.parser?.name !== "pdfjs" || reference.document?.parser?.version !== pdfjs.version) {
      throw new Error("PDF reference belongs to a different file, revision, or parser version.");
    }
    if (reference.text_content?.include_marked_content !== true || reference.text_content?.disable_normalization !== true) {
      throw new Error("PDF reference uses different text extraction options.");
    }
    if (reference.text_utf16_length !== undefined && (!Number.isSafeInteger(reference.text_utf16_length) || reference.text_utf16_length < 0)) {
      throw new Error("Invalid captured PDF text length.");
    }
    for (const end of [reference.start, reference.end]) {
      if (!end || !Number.isSafeInteger(end.page) || end.page < 1 || end.page > this.doc.numPages ||
          !Number.isSafeInteger(end.item_index) || end.item_index < 0 ||
          !Number.isSafeInteger(end.offset_utf16) || end.offset_utf16 < 0) throw new Error("Invalid PDF text endpoint.");
    }
    if (endpointOrder(reference.start, reference.end) >= 0) throw new Error("PDF text endpoints are reversed or empty.");
  }

  async read(reference, browserRange = null) {
    this.validate(reference);
    const anchors = [], pages = [], pageInfo = [], pageTexts = [];
    for (let number = reference.start.page; number <= reference.end.page; number++) {
      const data = await this.pageData(number);
      if (!this.alive) throw new Error("PDF document was closed during resolution.");
      const { page, content, entries } = data;
      let text = "";
      const pageAnchors = [];
      for (const endpoint of [reference.start, reference.end]) {
        if (endpoint.page !== number) continue;
        const item = content.items[endpoint.item_index];
        if (!item || typeof item.str !== "string" || endpoint.offset_utf16 > item.str.length) throw new Error("PDF endpoint is outside its text item.");
      }
      const mapping = browserRange ? await this.mapping(number) : null;
      for (const entry of entries) {
        const { item, index } = entry;
        if (number === reference.start.page && index < reference.start.item_index ||
            number === reference.end.page && index > reference.end.item_index) continue;
        const start = number === reference.start.page && index === reference.start.item_index ? reference.start.offset_utf16 : 0;
        const end = number === reference.end.page && index === reference.end.item_index ? reference.end.offset_utf16 : item.str.length;
        const lastItem = number === reference.end.page && index === reference.end.item_index;
        text += item.str.slice(start, end);
        if (item.hasEOL && !lastItem) text += "\n";
        if (end <= start) continue;
        const div = mapping?.entries.find(e => e.index === index)?.div;
        const domRange = div && browserRange && div.isConnected ? subrange(div, start, end) : null;
        const selectedQuads = pdfQuads(mapping?.view, domRange);
        const quads = selectedQuads.length ? selectedQuads : [itemQuad(item, content.styles[item.fontName])];
        pageAnchors.push({ page: number, page_ref: page.ref, item_index: index,
          start_utf16: start, end_utf16: end, text: item.str.slice(start, end),
          stability: "document-and-parser-relative", font_name: item.fontName, dir: item.dir,
          has_eol: item.hasEOL, transform: item.transform, width: item.width, height: item.height,
          geometry: { space: "pdf-page-user-space", unit: "user-unit", user_unit: page.userUnit,
            precision: selectedQuads.length ? "selected-dom-range" : "whole-text-item", quads, bbox: boxOf(quads) },
          ...(entry.markedId ? { marked_content_id: entry.markedId } : {}),
          ...(data.semantics.has(entry.markedId) ? { semantics: data.semantics.get(entry.markedId) } : {}) });
      }
      anchors.push(...pageAnchors);
      pages.push(number);
      pageInfo.push({ page: number, page_ref: page.ref, view_box: page.view,
        intrinsic_rotation: page.rotate, user_unit: page.userUnit, tagged: data.tagged });
      pageTexts.push(text);
    }
    const text = pageTexts.join("\n\n");
    if (Number.isInteger(reference.text_utf16_length) && reference.text_utf16_length !== text.length) {
      throw new Error("Resolved text length differs from the captured reference.");
    }
    return { schema: "office.selection.v1", file: this.documentFile, format: "pdf", kind: "pdf",
      document: this.identity, reference, text, anchors, pages, page_info: pageInfo };
  }

  apply(zoom, query, rescale = true) {
    this.options.zoom = zoom;
    this.options.query = query;
    if (!this.doc || !this.host.dataset.pdfReady) {
      // The initial call follows pagesinit, before the ready marker is set.
      if (!this.viewer.pagesCount) return;
    }
    if (rescale && Math.abs(this.viewer.currentScale * 100 - zoom) > 0.51) this.viewer.currentScaleValue = String(zoom / 100);
    if (this.query !== query) {
      this.query = query;
      this.findAgain(false, "");
    }
  }

  findAgain(previous = false, type = "again") {
    this.bus.dispatch("find", { source: this, type, query: this.query || "", caseSensitive: false,
      entireWord: false, highlightAll: true, findPrevious: previous, matchDiacritics: false });
  }

  async documentInfo() {
    if (!this.doc) throw new Error("No active PDF document.");
    const [metadata, outline, permissions] = await Promise.all([
      this.doc.getMetadata(), this.doc.getOutline(), this.doc.getPermissions()]);
    return { schema: "office.pdf.document.v1", file: this.documentFile,
      document: this.identity, info: metadata.info, metadata: metadata.metadata?.getAll() || null,
      outline, permissions };
  }

  destroy() {
    this.alive = false;
    this.resizeObserver.disconnect();
    if (this.resizeFrame != null) cancelAnimationFrame(this.resizeFrame);
    this.resizeFrame = null;
    clearTimeout(this.selectionTimer);
    this.selectionRevision++;
    this.viewer.setDocument(null);
    this.links.setDocument(null);
    this.abort.abort();
    this.task?.destroy().catch(() => {});
    this.pages.clear();
    delete this.host.dataset.pdfReady;
    delete this.host.dataset.pdfPainted;
    if (sessions.get(this.host.id) === this) sessions.delete(this.host.id);
  }
}

export function mount(host, options) {
  const session = new PdfSession(host, options);
  sessions.set(host.id, session);
  session.open(options.bytes).catch(error => {
    if (session.alive) session.state({ status: "error", message: error.message || String(error) });
  });
  return session;
}

export async function resolveSelection(surfaceId, value) {
  const session = sessions.get(surfaceId);
  if (!session) throw new Error("No active PDF surface.");
  const envelope = typeof value === "string" ? JSON.parse(value) : value;
  const reference = envelope.source?.reference || envelope.reference || envelope;
  return JSON.stringify(await session.read(reference), null, 2);
}

export async function documentInfo(surfaceId) {
  const session = sessions.get(surfaceId);
  if (!session) throw new Error("No active PDF surface.");
  return JSON.stringify(await session.documentInfo(), null, 2);
}
