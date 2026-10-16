// Host-neutral page thumbnails for Word and PDF. The live document remains
// in its original subtree and thumbnails never join document selections.
const CACHE_SIZE = 24;
const PARALLEL_RENDERS = 2;
let svgSequence = 0;

function wordPreview(svg) {
  const clone = svg.cloneNode(true);
  const box = svg.viewBox.baseVal;
  const width = box.width || Number.parseFloat(svg.getAttribute("width")) || 595;
  const height = box.height || Number.parseFloat(svg.getAttribute("height")) || 842;
  if (!box.width) clone.setAttribute("viewBox", `0 0 ${width} ${height}`);
  clone.setAttribute("width", "100%");
  clone.setAttribute("height", "100%");
  clone.setAttribute("aria-hidden", "true");
  clone.setAttribute("focusable", "false");
  // The rail is outside the document's system-font CSS scope. Preserve the
  // chosen glyph family without inheriting the main page's zoom or hit marks.
  const families = new Map(), texts = [...svg.querySelectorAll("text")];
  for (const [index, text] of [...clone.querySelectorAll("text")].entries()) {
    const family = texts[index].getAttribute("font-family") || "";
    if (!families.has(family)) families.set(family, getComputedStyle(texts[index]).fontFamily);
    text.style.fontFamily = families.get(family);
    text.classList.remove("office-search-hit");
  }
  // SVG definitions must not collide with the live page or another thumbnail.
  const ids = new Map(), prefix = `office-thumbnail-${++svgSequence}-`;
  for (const node of [clone, ...clone.querySelectorAll("[id]")]) {
    if (!node.id) continue;
    ids.set(node.id, prefix + node.id);
    node.id = prefix + node.id;
  }
  for (const node of [clone, ...clone.querySelectorAll("*")]) {
    for (const attr of [...node.attributes]) {
      const value = attr.value.replace(/url\(#([^)]*)\)/g, (match, id) =>
        ids.has(id) ? `url(#${ids.get(id)})` : match);
      const reference = value.startsWith("#") && ids.has(value.slice(1));
      if (value !== attr.value || reference) {
        attr.value = reference ? `#${ids.get(value.slice(1))}` : value;
      }
    }
  }
  return { node: clone, width, height };
}

class PageThumbnails {
  constructor(host, target, options, source, session) {
    Object.assign(this, { host, target, source, session, key: options.key, kind: options.kind });
    this.body = target;
    this.pages = this.kind === "word" ? [...source.querySelectorAll(".word-page")] : null;
    this.count = this.pages?.length ?? session.doc.numPages;
    this.active = session?.viewer.currentPageNumber || 1;
    this.entries = [];
    this.visible = new Set();
    this.cache = new Map();
    this.queue = [];
    this.queued = new Set();
    this.running = 0;
    this.alive = true;
    this.open = false;
    this.abort = new AbortController();
    this.scroller = this.kind === "word" ? host : host.querySelector(".pdf-container");
    this.onPdfPage = event => this.setActive(event.pageNumber, true);
    this.onPdfRotation = () => this.invalidate();
    this.intersections = new IntersectionObserver(entries => {
      for (const { target, isIntersecting } of entries) {
        const entry = this.entries[Number(target.dataset.page) - 1];
        if (!entry) continue;
        if (isIntersecting) this.visible.add(entry);
        else this.visible.delete(entry);
      }
      this.scheduleVisible();
    }, { root: this.body, rootMargin: "180px 0px" });
    this.resize = new ResizeObserver(() => {
      this.schedulePosition();
    });
    this.labels = new MutationObserver(() => this.updateLabels());
  }

  current() {
    return this.alive && this.host.isConnected && this.target.isConnected &&
      this.host.__officeViewerPageThumbnails === this &&
      (this.kind === "word" ? this.host.querySelector(".word-pages") === this.source :
        this.host.__officeViewerPdf === this.session && this.session.alive);
  }

  start(options) {
    const { signal } = this.abort;
    this.body.addEventListener("click", event => {
      const button = event.target.closest("button[data-page]");
      if (button && this.body.contains(button)) this.navigate(Number(button.dataset.page));
    }, { signal });
    this.body.addEventListener("keydown", event => {
      const button = event.target.closest("button[data-page]");
      if (!button) return;
      let page = Number(button.dataset.page);
      if (event.key === "ArrowDown" || event.key === "ArrowRight") page++;
      else if (event.key === "ArrowUp" || event.key === "ArrowLeft") page--;
      else if (event.key === "Home") page = 1;
      else if (event.key === "End") page = this.count;
      else return;
      event.preventDefault();
      page = Math.max(1, Math.min(this.count, page));
      this.entries[page - 1]?.button.focus({ preventScroll: true });
      this.navigate(page);
    }, { signal });
    this.scroller.addEventListener("scroll", () => this.schedulePosition(), { passive: true, signal });
    this.resize.observe(this.host);
    this.labels.observe(this.target, { attributes: true, attributeFilter: ["data-page-label", "data-unavailable-label"] });
    this.session?.bus.on("pagechanging", this.onPdfPage);
    this.session?.bus.on("rotationchanging", this.onPdfRotation);
    this.update(options);
  }

  updateLabels() {
    for (const entry of this.entries) {
      entry.button.setAttribute("aria-label", this.target.dataset.pageLabel.replace("{page}", entry.page));
      if (entry.failed) entry.paper.textContent = this.target.dataset.unavailableLabel;
    }
  }

  makeEntries() {
    if (this.entries.length || !this.count) return;
    const fragment = document.createDocumentFragment();
    for (let page = 1; page <= this.count; page++) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "office-page-item";
      button.dataset.page = page;
      button.tabIndex = page === this.active ? 0 : -1;
      if (page === this.active) button.setAttribute("aria-current", "page");
      const number = document.createElement("span");
      number.className = "office-page-number";
      number.textContent = page;
      number.setAttribute("aria-hidden", "true");
      const paper = document.createElement("span");
      paper.className = "office-page-paper";
      paper.setAttribute("aria-hidden", "true");
      button.append(number, paper);
      const entry = { page, button, paper, version: 0, failed: false };
      this.entries.push(entry);
      fragment.append(button);
    }
    this.body.replaceChildren(fragment);
    this.updateLabels();
  }

  update(options) {
    const changedFonts = this.systemFonts !== options.systemFonts;
    this.systemFonts = options.systemFonts;
    if (changedFonts && this.kind === "word") this.invalidate();
    if (this.open !== options.enabled) {
      this.open = options.enabled;
      if (this.open) {
        this.makeEntries();
        for (const entry of this.entries) this.intersections.observe(entry.button);
        this.revealActive();
      } else {
        this.intersections.disconnect();
        this.visible.clear();
        this.cancelPending();
      }
    }
    this.schedulePosition();
  }

  setActive(page, reveal = false) {
    if (!this.current() || !Number.isInteger(page) || page < 1 || page > this.count) return;
    if (this.active !== page) {
      const old = this.entries[this.active - 1]?.button;
      old?.removeAttribute("aria-current");
      if (old) old.tabIndex = -1;
      this.active = page;
      const button = this.entries[page - 1]?.button;
      button?.setAttribute("aria-current", "page");
      if (button) button.tabIndex = 0;
    }
    if (reveal && this.open) this.revealActive();
  }

  revealActive() {
    const button = this.entries[this.active - 1]?.button;
    if (!button) return;
    const view = this.body.getBoundingClientRect(), rect = button.getBoundingClientRect();
    if (rect.top < view.top || rect.bottom > view.bottom) {
      this.body.scrollTop += rect.top - view.top - Math.max(0, (view.height - rect.height) / 2);
    }
  }

  navigate(page) {
    if (!this.current() || page < 1 || page > this.count) return;
    this.setActive(page, true);
    if (this.kind === "pdf") this.session.viewer.scrollPageIntoView({ pageNumber: page });
    else {
      const rect = this.pages[page - 1].getBoundingClientRect();
      const view = this.host.getBoundingClientRect();
      this.host.scrollTop += rect.top - view.top - 20;
      this.schedulePosition();
    }
  }

  schedulePosition() {
    if (!this.current() || this.positionFrame) return;
    this.positionFrame = requestAnimationFrame(() => {
      this.positionFrame = null;
      if (!this.current() || this.kind !== "word" || !this.pages.length) return;
      const view = this.host.getBoundingClientRect();
      // Find nearby pages without reading every page's geometry on scroll.
      let low = 0, high = this.pages.length;
      while (low < high) {
        const middle = (low + high) >> 1;
        if (this.pages[middle].getBoundingClientRect().top < view.top) low = middle + 1;
        else high = middle;
      }
      const score = index => {
        const rect = this.pages[index].getBoundingClientRect();
        const visible = Math.max(0, Math.min(view.bottom, rect.bottom) - Math.max(view.top, rect.top));
        return visible / Math.max(1, Math.min(rect.height, view.height));
      };
      let best = this.active - 1, bestScore = score(best);
      for (let i = Math.max(0, low - 1); i < Math.min(this.pages.length, low + 8); i++) {
        const value = score(i);
        if (value > bestScore + 0.02) { best = i; bestScore = value; }
      }
      this.setActive(best + 1, true);
    });
  }

  scheduleVisible() {
    if (!this.current() || !this.open) return;
    for (const entry of this.visible) {
      if (this.cache.has(entry.page)) {
        this.cache.delete(entry.page);
        this.cache.set(entry.page, entry);
      } else if (!entry.task && !entry.failed && !this.queued.has(entry)) {
        this.queue.push(entry);
        this.queued.add(entry);
      }
    }
    this.pump();
  }

  pump() {
    while (this.current() && this.open && this.running < PARALLEL_RENDERS && this.queue.length) {
      const entry = this.queue.shift();
      this.queued.delete(entry);
      if (!this.visible.has(entry) || this.cache.has(entry.page)) continue;
      this.running++;
      const version = entry.version;
      entry.task = true;
      this.render(entry, version).catch(error => {
        if (this.current() && entry.version === version && error.name !== "RenderingCancelledException") {
          entry.failed = true;
          entry.paper.textContent = this.target.dataset.unavailableLabel;
        }
      }).finally(() => {
        if (entry.version === version) entry.task = null;
        this.running--;
        this.pump();
      });
    }
  }

  async render(entry, version) {
    let preview;
    if (this.kind === "word") {
      preview = wordPreview(this.pages[entry.page - 1].querySelector("svg"));
    } else {
      const page = await this.session.doc.getPage(entry.page);
      if (!this.current() || entry.version !== version || !this.open) return;
      const rotation = (page.rotate + this.session.viewer.pagesRotation) % 360;
      const base = page.getViewport({ scale: 1, rotation });
      const width = Math.max(1, entry.paper.clientWidth || 132);
      const ratio = Math.min(2, window.devicePixelRatio || 1);
      const viewport = page.getViewport({ scale: width * ratio / base.width, rotation });
      const canvas = document.createElement("canvas");
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const task = entry.task = page.render({ canvasContext: canvas.getContext("2d"), viewport });
      await task.promise;
      preview = { node: canvas, width: base.width, height: base.height };
    }
    if (!this.current() || !this.open || entry.version !== version) return;
    entry.paper.style.aspectRatio = `${preview.width} / ${preview.height}`;
    entry.paper.replaceChildren(preview.node);
    entry.button.dataset.painted = "true";
    this.cache.set(entry.page, entry);
    this.evict();
  }

  release(entry) {
    const canvas = entry.paper.querySelector("canvas");
    if (canvas) canvas.width = canvas.height = 0;
    entry.paper.replaceChildren();
    delete entry.button.dataset.painted;
  }

  evict() {
    for (const [page, entry] of this.cache) {
      if (this.cache.size <= CACHE_SIZE) break;
      if (!this.visible.has(entry)) { this.release(entry); this.cache.delete(page); }
    }
  }

  cancelPending() {
    this.queue.length = 0;
    this.queued.clear();
    for (const entry of this.entries) {
      entry.version++;
      if (entry.task?.cancel) entry.task.cancel();
      entry.task = null;
    }
  }

  invalidate() {
    this.cancelPending();
    for (const entry of this.entries) { this.release(entry); entry.failed = false; }
    this.cache.clear();
    this.scheduleVisible();
  }

  dispose() {
    if (!this.alive) return;
    this.alive = false;
    this.cancelPending();
    cancelAnimationFrame(this.positionFrame);
    this.intersections.disconnect();
    this.resize.disconnect();
    this.labels.disconnect();
    this.abort.abort();
    this.session?.bus.off("pagechanging", this.onPdfPage);
    this.session?.bus.off("rotationchanging", this.onPdfRotation);
    for (const entry of this.entries) this.release(entry);
    this.cache.clear();
    this.body.replaceChildren();
    if (this.host.__officeViewerPageThumbnails === this) delete this.host.__officeViewerPageThumbnails;
  }
}

export function mount(host, target, options) {
  const source = options.kind === "word" ? host.querySelector(".word-pages") : null;
  const session = options.kind === "pdf" ? host.__officeViewerPdf : null;
  const current = host.__officeViewerPageThumbnails;
  if (current?.key === options.key && current.target === target && current.source === source &&
      current.session === session && current.current()) {
    current.update(options);
    return;
  }
  current?.dispose();
  if (!source && (!session?.alive || !session.doc)) return;
  const manager = new PageThumbnails(host, target, options, source, session);
  host.__officeViewerPageThumbnails = manager;
  manager.start(options);
}
