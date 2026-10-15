import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { before, after, test } from 'node:test';
import { runtime } from './browser-runtime.mjs';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const styles = await readFile(new URL('../public/styles.css', import.meta.url), 'utf8');
let browser;
before(async () => {
  browser = await chromium.launch({ headless: true,
    ...(process.env.CHROME_EXECUTABLE ? { executablePath: process.env.CHROME_EXECUTABLE } : {}),
  });
});
after(async () => { await browser?.close(); });

const identity = { file: '/docs/review.docx', sha256: 'snapshot-a' };
const comment = (id, options = {}) => ({
  id, type: 'docx.comment', document: identity, author: 'Ada Lovelace',
  text: 'A😀B\nSecond line', text_status: 'available', parent_id: null,
  anchor: { kind: 'docx', locations: [{ part: 'word/document.xml', path: 'p[1]' }] }, ...options,
});
const word = (body, items = []) => `<div class="word-pages"><div class="word-page" data-page="1">
  <svg width="640" height="900" xmlns="http://www.w3.org/2000/svg">${body}</svg></div></div>
  <script type="application/json" class="office-annotations-metadata">${JSON.stringify({ schema: 'office.annotations.v1', items })}</script>`;
const paragraph = (value, attrs = '') => `<text x="40" y="80" font-size="16" data-source-part="word/document.xml" data-source-path="p[1]" ${attrs}>${value}</text>`;

async function setup(t, html, kind = 'word', options = {}) {
  const page = await browser.newPage({ viewport: { width: 1500, height: 1100 } });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  t.after(async () => { await page.close(); assert.deepEqual(errors, [], 'no browser runtime errors'); });
  await page.setContent(`<style>${styles} #document {width:1250px;margin:20px}</style>
    <div id="document" class="surface office-viewer-surface ${kind === 'word' ? 'word-surface' : 'sheet-surface'}"></div>`);
  if (options.noObserver) await page.evaluate(() => { window.ResizeObserver = undefined; });
  if (options.trackListeners) await page.evaluate(() => {
    window.runtimeListeners = [];
    window.trackingRuntime = false;
    const add = EventTarget.prototype.addEventListener;
    const remove = EventTarget.prototype.removeEventListener;
    EventTarget.prototype.addEventListener = function(type, handler, options) {
      if (window.trackingRuntime && (this === window || this === document)) window.runtimeListeners.push({target:this,type,handler});
      return add.call(this,type,handler,options);
    };
    EventTarget.prototype.removeEventListener = function(type, handler, options) {
      const index = window.runtimeListeners.findIndex(entry => entry.target === this && entry.type === type && entry.handler === handler);
      if (index >= 0) window.runtimeListeners.splice(index,1);
      return remove.call(this,type,handler,options);
    };
  });
  await page.addScriptTag({ content: runtime });
  await page.evaluate(({ html, kind }) => {
    window.events = []; window.annotations = [];
    window.installFixture = (content = html, file = kind === 'word' ? '/docs/review.docx' : '/docs/review.xlsx') => {
      window.trackingRuntime = true;
      try {
        window.officeTest.install('document', kind, file, content,
          value => window.events.push(JSON.parse(value)),
          (item, start, end, scope) => window.annotations.push({ item: JSON.parse(item), start, end, scope }));
      } finally { window.trackingRuntime = false; }
    };
    window.installFixture();
  }, { html, kind });
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  return page;
}

async function selectText(page, startSelector, start, endSelector, end) {
  await page.evaluate(({ startSelector, start, endSelector, end }) => {
    const range = document.createRange();
    range.setStart(document.querySelector(startSelector).firstChild, start);
    range.setEnd(document.querySelector(endSelector).firstChild, end);
    const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));
  }, { startSelector, start, endSelector, end });
}

test('native Word selection preserves UTF-16 logical anchors and physical fallback fragments', async t => {
  const html = word(paragraph('A😀BC', `id="joined" data-source-id="paragraph-1" data-source-selector="/docx/body/p[1]"
    data-source-start="10" data-source-end="15" data-source-runs="r[1],r[2]" data-source-actionable="true"
    data-source-para-id="ABCD" data-source-anchor-status="stable"`) +
    '<text id="physical" x="40" y="120">second line</text>');
  const page = await setup(t, html);
  await selectText(page, '#joined', 1, '#physical', 6);
  const event = await page.evaluate(() => window.events.at(-1));
  assert.equal(event.schema, 'office.selection.v1');
  assert.equal(event.file, identity.file); assert.equal(event.format, 'docx'); assert.equal(event.kind, 'word');
  assert.deepEqual(event.pages, [1]);
  assert.equal(event.anchors.length, 2);
  const [joined, physical] = event.anchors;
  assert.equal(joined.text, '😀BC'); assert.equal(joined.start_utf16, 11); assert.equal(joined.end_utf16, 15);
  assert.equal(joined.selector, '/docx/body/p[1]'); assert.equal(joined.source_id, 'paragraph-1');
  assert.equal(joined.actionable, true); assert.equal(joined.para_id, 'ABCD');
  assert.deepEqual(joined.runs, ['r[1]', 'r[2]']);
  assert.equal(physical.text, 'second'); assert.equal(physical.stability, 'physical-only');
  assert.equal('selector' in physical, false); assert.equal('path' in physical, false);
});

test('incomplete or mismatched Word source spans never become logical anchors', async t => {
  for (const attrs of ['data-source-end="5"', 'data-source-start="0" data-source-end="999"']) {
    const page = await setup(t, word(paragraph('Hello', `id="source" data-source-selector="/docx/body/p[1]" ${attrs}`)));
    await selectText(page, '#source', 1, '#source', 4);
    const event = await page.evaluate(() => window.events.at(-1));
    assert.equal(event.anchors[0].stability, 'physical-only');
    assert.equal(event.anchors[0].text, 'ell');
    assert.equal('selector' in event.anchors[0], false);
  }
});

function excel() {
  return `<div class="sheet-document"><section data-sheet="First"><h2>First</h2>
    <table><colgroup>${'<col style="width:100px">'.repeat(4)}</colgroup><tbody>
    <tr><td rowspan="2" colspan="2">Merged</td><td>10</td><td>11</td></tr>
    <tr><td>20</td><td>21</td></tr><tr><td>30</td><td>31</td><td>32</td><td>33</td></tr>
    </tbody></table></section><section data-sheet="Second"><h2>Second</h2><table><tbody><tr><td>Other</td></tr></tbody></table></section></div>
    <script type="application/json" class="office-formula-metadata">${JSON.stringify({sheets:{First:{C1:'SUM(A3:B3)'}}})}</script>`;
}

test('Excel drag selection accounts for merged cells and produces canonical source references', async t => {
  const page = await setup(t, excel(), 'excel');
  const start = await page.locator('section[data-sheet="First"] td[data-cell="A1"]').boundingBox();
  const end = await page.locator('section[data-sheet="First"] td[data-cell="D2"]').boundingBox();
  await page.mouse.move(start.x + 8, start.y + 8); await page.mouse.down();
  await page.mouse.move(end.x + end.width / 2, end.y + end.height / 2, { steps: 4 }); await page.mouse.up();
  const event = await page.evaluate(() => window.events.at(-1));
  assert.equal(event.range, 'A1:D2'); assert.equal(event.sheet, 'First');
  assert.equal(event.selector, '/xlsx/sheet[name="First"]/range[A1:D2]');
  assert.deepEqual(event.cells.map(cell => cell.address), ['A1', 'C1', 'D1', 'C2', 'D2']);
  assert.deepEqual(event.cells[0], {address:'A1',row:1,column:1,text:'Merged'});
  assert.equal(await page.locator('[data-excel-name]').textContent(), 'A1:D2');
  assert.equal(await page.locator('[data-excel-value]').textContent(), '21');
});

test('Excel formula inspection, sheet switching, cancellation, and explicit clearing preserve state', async t => {
  const page = await setup(t, excel(), 'excel');
  await page.locator('td[data-cell="C1"]').click();
  assert.equal(await page.locator('[data-excel-value]').textContent(), '=SUM(A3:B3)');
  await page.getByRole('tab', {name:'Second',exact:true}).click();
  assert.equal(await page.locator('section[data-sheet="First"]').evaluate(node => node.hidden), true);
  assert.equal(await page.locator('td.office-cell-selected').count(), 0);
  await page.locator('section[data-sheet="Second"] td').click();
  assert.equal((await page.evaluate(() => window.events.at(-1))).sheet, 'Second');
  await page.evaluate(() => window.officeTest.clear('document'));
  assert.equal(await page.locator('td.office-cell-selected, td.office-cell-active').count(), 0);
  const cell = page.locator('section[data-sheet="Second"] td');
  await cell.dispatchEvent('pointerdown', {button:0});
  await page.evaluate(() => window.dispatchEvent(new Event('pointercancel')));
  assert.equal(await page.locator('td.office-cell-selected').count(), 0);
});

test('annotation requests resolve the installed snapshot and reject stale, duplicate, and malformed identities', async t => {
  const original = comment('c1');
  const html = word(paragraph('Source'), [original]);
  const page = await setup(t, html);
  const results = await page.evaluate(original => {
    const select = item => window.officeTest.select('document', JSON.stringify(item));
    const stale = select({...original, document:{...original.document,sha256:'old-snapshot'}});
    const forged = select({...original,text:'injected',author:'attacker',anchor:{}});
    const invalid = document.getElementById('document').__officeSelectAnnotation(undefined);
    return {stale,forged,invalid,events:window.annotations};
  }, original);
  assert.equal(results.stale, false); assert.equal(results.invalid, false); assert.equal(results.forged, true);
  assert.deepEqual(results.events[0].item, original);
  await page.evaluate(html => window.installFixture(html), word(paragraph('Source'), [original, {...original,text:'duplicate'}]));
  assert.equal(await page.evaluate(original => window.officeTest.select('document', JSON.stringify(original)), original), false);
});

test('annotation body ranges retain exact newlines and reject split surrogate pairs', async t => {
  const page = await setup(t, word(paragraph('Source'), [comment('c1')]));
  await selectText(page, '.office-annotation-body', 1, '.office-annotation-body', 4);
  let events = await page.evaluate(() => window.annotations);
  assert.equal(events.at(-1).start, 1); assert.equal(events.at(-1).end, 4);
  assert.equal(events.at(-1).item.text.slice(1,4), '😀B');
  const count = events.length;
  await selectText(page, '.office-annotation-body', 2, '.office-annotation-body', 4);
  events = await page.evaluate(() => window.annotations);
  assert.equal(events.length, count);
  await page.evaluate(() => window.officeTest.select('document', JSON.stringify({id:'c1',document:{file:'/docs/review.docx',sha256:'snapshot-a'}})));
  events = await page.evaluate(() => window.annotations);
  assert.equal(events.at(-1).end, 'A😀B\nSecond line'.length);
  assert.deepEqual(await page.evaluate(() => window.events), []);
});

test('paginated comment lists retain source order, expand, and tolerate reply cycles', async t => {
  const items = Array.from({length:201},(_,i) => comment(`c${i}`, {anchor:{kind:'docx',locations:[]},text:`Comment ${i}`}));
  const page = await setup(t, word(paragraph('Source'), items));
  assert.equal(await page.locator('.office-context-card').count(),200);
  await page.getByRole('button',{name:'Show 1 more',exact:true}).click();
  assert.equal(await page.locator('.office-context-card').count(),201);
  assert.equal(await page.locator('.office-context-entry').last().getAttribute('data-office-annotation-id'),'c200');
  await page.evaluate(html => window.installFixture(html),word(paragraph('Source'),[
    comment('parent',{parent_id:'reply'}),comment('reply',{parent_id:'parent'}),
  ]));
  assert.equal(await page.locator('.office-context-card').count(),2);
});

test('zoom and search work without ResizeObserver and disposed surfaces stop dispatching', async t => {
  const page = await setup(t, word(paragraph('Search text', 'id="source"'),[comment('c1')]),'word',{noObserver:true});
  await page.evaluate(() => { window.officeTest.zoom('document',1.5); window.officeTest.search('document','SEARCH'); });
  assert.equal(await page.locator('#source.office-search-hit').count(),1);
  assert.equal(await page.locator('#document').evaluate(node => node.style.getPropertyValue('--office-zoom')),'1.5');
  await page.evaluate(() => window.officeTest.search('document',''));
  assert.equal(await page.locator('.office-search-hit').count(),0);
  await page.evaluate(() => window.officeTest.dispose('document'));
  await selectText(page,'#source',0,'#source',6);
  assert.deepEqual(await page.evaluate(() => window.events),[]);
  assert.equal(await page.locator('.office-context-rail').count(),0);
  assert.equal(await page.evaluate(identity => window.officeTest.focus('document',JSON.stringify({id:'c1',document:identity})),identity),false);
});


test('replacing or detaching the root retires the old snapshot and global listeners', async t => {
  const page = await setup(t, word(paragraph('Source'),[comment('c1')]), 'word', {trackListeners:true});
  const original = await page.evaluate(() => window.runtimeListeners.length);
  assert.ok(original > 0);
  await page.evaluate(() => {
    const host = document.getElementById('document');
    host.replaceWith(host.cloneNode(false));
    window.installFixture();
  });
  assert.equal(await page.evaluate(() => window.runtimeListeners.length), original);
  await page.evaluate(() => {
    document.getElementById('document').remove();
    window.officeTest.dispose('document');
  });
  assert.equal(await page.evaluate(() => window.runtimeListeners.length), 0);
});

async function cellCenter(page, selector) {
  const box = await page.locator(selector).boundingBox();
  assert.ok(box, `visible drag target: ${selector}`);
  return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
}

async function startExcelRange(page) {
  const start = await cellCenter(page, 'section[data-sheet="First"] td[data-cell="A3"]');
  const end = await cellCenter(page, 'section[data-sheet="First"] td[data-cell="B3"]');
  await page.mouse.move(start.x, start.y);
  await page.mouse.down();
  await page.mouse.move(end.x, end.y);
  return end;
}

test('Excel release over unrelated or malformed cells reports the last valid highlighted range', async t => {
  for (const invalid of ['other-table', 'missing-address', 'missing-row', 'missing-column', 'invalid-coordinate']) {
    const page = await setup(t, excel(), 'excel');
    await page.evaluate(invalid => {
      let cell;
      if (invalid === 'other-table') {
        const table = document.createElement('table');
        table.style.cssText = 'position:fixed;left:1300px;top:800px;width:150px;z-index:9999;background:white';
        table.innerHTML = '<tbody><tr><td data-cell="Z9" data-row="9" data-col="26">Unrelated</td></tr></tbody>';
        document.body.appendChild(table);
        cell = table.querySelector('td');
      } else {
        cell = document.querySelector('section[data-sheet="First"] td[data-cell="C3"]');
        if (invalid === 'missing-address') cell.removeAttribute('data-cell');
        if (invalid === 'missing-row') cell.removeAttribute('data-row');
        if (invalid === 'missing-column') cell.removeAttribute('data-col');
        if (invalid === 'invalid-coordinate') cell.dataset.row = 'NaN';
      }
      cell.id = 'invalid-drag-target';
    }, invalid);
    await startExcelRange(page);
    const target = await cellCenter(page, '#invalid-drag-target');
    assert.equal(await page.evaluate(point => document.elementFromPoint(point.x, point.y)?.closest('td')?.id, target), 'invalid-drag-target', `${invalid}: drag hits the invalid cell`);
    await page.mouse.move(target.x, target.y);
    await page.mouse.up();
    const result = await page.evaluate(() => ({
      events: window.events,
      highlighted: [...document.querySelectorAll('td.office-cell-selected')].map(cell => cell.dataset.cell),
      active: [...document.querySelectorAll('td.office-cell-active')].map(cell => cell.dataset.cell),
    }));
    assert.equal(result.events.length, 1, `${invalid}: highlighted range reaches the consumer`);
    assert.equal(result.events[0].range, 'A3:B3', invalid);
    assert.deepEqual(result.events[0].cells.map(cell => cell.address), ['A3', 'B3'], invalid);
    assert.deepEqual(result.highlighted, ['A3', 'B3'], invalid);
    assert.deepEqual(result.active, ['B3'], invalid);
  }
});

test('Excel movement within one cell or outside the grid skips full-table scans', async t => {
  const page = await setup(t, excel(), 'excel');
  const active = await startExcelRange(page);
  await page.evaluate(() => {
    window.tableScans = 0;
    const table = document.querySelector('section[data-sheet="First"] table');
    const query = table.querySelectorAll;
    table.querySelectorAll = function(selector) {
      if (selector === 'td[data-cell]' || selector === 'td.office-cell-selected, td.office-cell-active') window.tableScans++;
      return query.call(this, selector);
    };
  });
  for (let i = 1; i <= 5; i++) await page.mouse.move(active.x + i, active.y);
  assert.equal(await page.evaluate(() => window.tableScans), 0, 'same-cell movement does not rescan');
  await page.mouse.move(1450, 1000);
  await page.evaluate(() => window.dispatchEvent(new Event('pointermove')));
  assert.equal(await page.evaluate(() => window.tableScans), 0, 'outside-grid and non-mouse movement do not rescan');
  await page.mouse.up();
  const events = await page.evaluate(() => window.events);
  assert.equal(events.length, 1);
  assert.equal(events[0].range, 'A3:B3');
});
