import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { after, before, test } from 'node:test';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
import { runtime } from './browser-runtime.mjs';
const styles = await readFile(new URL('../public/styles.css', import.meta.url), 'utf8');
let browser;
before(async () => {
  browser = await chromium.launch({ headless: true,
    ...(process.env.CHROME_EXECUTABLE ? { executablePath: process.env.CHROME_EXECUTABLE } : {}),
  });
});
after(async () => { await browser?.close(); });

const identity = { file: '/docs/review.docx', sha256: 'fixture-snapshot' };
const items = [
  { author: '产品评审', initials: 'PR' },
  { author: 'Ada Lovelace', initials: null },
  { author: null, initials: null },
].map((person, index) => ({
  ...person, id: `docx.comment:${index}`, type: 'docx.comment', document: identity,
  text: `Review comment ${index + 1}: preserve this complete comment body.`,
  text_status: 'available', resolved: null, parent_id: null, date: null,
  anchor: { kind: 'docx', locations: [{ part: 'word/document.xml', path: `p[${index + 1}]` }] },
}));
const html = `<div class="word-pages"><div class="word-page" data-page="1">
  <svg width="640" height="900" xmlns="http://www.w3.org/2000/svg">
    ${items.map((_, i) => `<text x="40" y="${120 + i * 35}" font-size="16"
      data-source-part="word/document.xml" data-source-path="p[${i + 1}]">Source paragraph ${i + 1}</text>`).join('')}
  </svg></div></div><script type="application/json" class="office-annotations-metadata">${JSON.stringify({ schema: 'office.annotations.v1', items })}</script>`;

async function settle(page) {
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
}
async function setup(t, width = 1250, viewportWidth = 1500, content = html, kind = 'word') {
  const page = await browser.newPage({ viewport: { width: viewportWidth, height: 1100 } });
  t.after(() => page.close());
  await page.setContent(`<style>${styles}\n#document { width:${width}px; margin:20px; }</style>
    <div id="document" class="surface office-viewer-surface ${kind === 'word' ? 'word-surface' : 'sheet-surface'}"></div>`);
  await page.addScriptTag({ content: runtime });
  await page.evaluate(({ html, file, kind }) => {
    window.annotationEvents = [];
    window.documentEvents = [];
    window.installFixture = () => window.officeTest.install('document', kind, file, html,
      event => window.documentEvents.push(event),
      (item, start, end, scope) => window.annotationEvents.push({ item: JSON.parse(item), start, end, scope }));
    window.installFixture();
  }, { html: content, file: kind === 'word' ? identity.file : '/docs/review.xlsx', kind });
  await settle(page);
  return page;
}
async function geometry(page) {
  return page.locator('#document').evaluate(host => {
    const paper = host.querySelector('.word-page').getBoundingClientRect();
    const rail = host.querySelector('.office-context-rail').getBoundingClientRect();
    const cards = [...host.querySelectorAll('.office-context-placed > .office-context-card')]
      .map(node => { const box = node.getBoundingClientRect(); return { left: box.left, top: box.top, bottom: box.bottom }; });
    const source = host.querySelector('svg text[data-source-path]').getBoundingClientRect();
    return { gap: cards[0].left - paper.right, railWidth: rail.width,
      sourceTop: source.top, cards, pageWidth: paper.width };
  });
}

test('avatars stay inside their own author rows when nearby comments stack', async t => {
  const page = await setup(t);
  const avatars = page.locator('.office-annotation-avatar');
  assert.equal(await avatars.count(), 3);
  assert.equal(await page.locator('.office-context-person > .office-annotation-avatar').count(), 3);
  assert.equal(await avatars.nth(0).textContent(), 'PR');
  assert.equal(await avatars.nth(1).textContent(), 'AL');
  assert.equal(await avatars.nth(2).locator('svg').count(), 1);
  assert.equal(await page.locator('.office-context-author').nth(2).textContent(), 'Unknown author');
  assert.equal(await page.locator('.office-comment-marker:visible').count(), 0);
  const layout = await geometry(page);
  assert.ok(layout.cards[1].top > layout.sourceTop + 100, 'nearby paragraphs force card collision avoidance');
  for (let i = 1; i < layout.cards.length; i++) assert.ok(layout.cards[i].top >= layout.cards[i - 1].bottom);
});

test('paper and comment cards remain 12px apart after zoom and resizing', async t => {
  const page = await setup(t);
  let layout = await geometry(page);
  assert.ok(Math.abs(layout.gap - 12) < 1);
  assert.ok(Math.abs(layout.cards[0].top - layout.sourceTop) < 1);
  await page.locator('#document').evaluate(host => {
    host.style.setProperty('--office-zoom', '1.25');
    host.style.width = '1400px';
    host.__officeLayoutAnnotations();
  });
  await settle(page);
  layout = await geometry(page);
  assert.ok(Math.abs(layout.gap - 12) < 1);
  assert.ok(Math.abs(layout.pageWidth - 800) < 1);
  assert.ok(Math.abs(layout.cards[0].top - layout.sourceTop) < 1);
});

test('collapsed source bubbles open the relevant card without quoting a body', async t => {
  const page = await setup(t, 900);
  const rail = page.locator('.office-context-rail');
  assert.equal(await rail.evaluate(node => node.open), false);
  assert.equal(Math.round((await geometry(page)).railWidth), 36);
  assert.equal(await page.locator('.office-comment-marker:visible').count(), 3);
  assert.equal(await page.locator('.office-comment-marker .office-annotation-avatar').count(), 0);
  await page.locator('.office-comment-marker').nth(1).click();
  await settle(page);
  assert.equal(await rail.evaluate(node => node.open), true);
  assert.equal(await page.locator('.office-comment-marker:visible').count(), 0);
  assert.equal(await page.locator('.office-context-card').nth(1).evaluate(node => node.classList.contains('office-context-active')), true);
  assert.equal(await page.locator('[data-source-path="p[2]"].office-annotation-focus').count(), 1);
  assert.deepEqual(await page.evaluate(() => window.annotationEvents), []);
});

test('comment quoting and native body selections retain their source identity', async t => {
  const page = await setup(t);
  await page.locator('.office-context-quote-action').first().click();
  let events = await page.evaluate(() => window.annotationEvents);
  assert.equal(events[0].item.id, items[0].id);
  assert.deepEqual(events[0].item.document, identity);
  assert.equal(events[0].scope, 'whole');
  assert.equal(events[0].end, items[0].text.length);
  await page.locator('.office-annotation-body').nth(1).evaluate(node => {
    const range = document.createRange();
    range.setStart(node.firstChild, 0); range.setEnd(node.firstChild, 6);
    const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));
  });
  events = await page.evaluate(() => window.annotationEvents);
  const selected = events.at(-1);
  assert.equal(selected.item.id, items[1].id);
  assert.equal(selected.start, 0); assert.equal(selected.end, 6); assert.equal(selected.scope, 'text');
});

test('a distant source bubble opens and scrolls to its fully positioned card', async t => {
  const content = html.replace('height="900"', 'height="2200"')
    .replace('y="155"', 'y="820"').replace('y="190"', 'y="1520"');
  const page = await setup(t, 900, 1200, content);
  await page.setViewportSize({ width: 1200, height: 600 });
  await page.locator('.office-comment-marker').nth(2).click();
  await settle(page);
  const card = await page.locator('.office-context-active').boundingBox();
  assert.ok(card.y >= 0 && card.y + card.height <= 600, 'expanded distant card stays inside the viewport');
});

test('Excel Note markers preserve whole-body annotation selection', async t => {
  const note = { ...items[0], id: 'xlsx.note:0', type: 'xlsx.note',
    document: { ...identity, file: '/docs/review.xlsx' },
    anchor: { kind: 'xlsx', sheet: 'Sheet1', cell: 'A1' } };
  const content = `<div class="excel-workspace"><div class="sheet-document office-excel-workbook">
    <section data-sheet="Sheet1"><table><tbody><tr>
    <td data-sheet="Sheet1" data-cell="A1" data-row="1" data-col="1">A1</td>
    </tr></tbody></table></section></div></div>
    <script type="application/json" class="office-annotations-metadata">${JSON.stringify({ schema: 'office.annotations.v1', items: [note] })}</script>`;
  const page = await setup(t, 900, 1200, content, 'excel');
  await page.locator('.office-note-marker').click();
  const events = await page.evaluate(() => window.annotationEvents);
  assert.equal(events.length, 1);
  assert.equal(events[0].item.id, note.id);
  assert.equal(events[0].scope, 'whole');
  assert.equal(events[0].end, note.text.length);
});

test('reinstall and disposal leave no duplicate cards, avatars, or styles', async t => {
  const page = await setup(t);
  await page.evaluate(() => window.installFixture());
  await settle(page);
  assert.equal(await page.locator('.office-context-rail').count(), 1);
  assert.equal(await page.locator('.office-annotation-avatar').count(), 3);
  assert.equal(await page.locator('.office-annotation-avatar-styles').count(), 1);
  await page.locator('#document').evaluate(host => host.__officeViewerCleanup());
  assert.equal(await page.locator('.office-context-rail').count(), 0);
  assert.equal(await page.locator('.office-annotation-avatar-styles').count(), 0);
});

test('proven comment-context metadata supports embedded host source navigation', async t => {
  const content = html.replaceAll('data-source-part', 'data-comment-part')
    .replaceAll('data-source-path', 'data-comment-path');
  const page = await setup(t, 1250, 1500, content);
  assert.equal(await page.locator('.office-context-placed > .office-context-card').count(), 3);
  await page.locator('.office-context-card').nth(1).getByRole('button', { name: 'Show source' }).click();
  assert.equal(await page.locator('[data-comment-path="p[2]"].office-annotation-focus').count(), 1);
});

test('replies retain their own author avatar and quote identity inside the parent card', async t => {
  const replies = items.map((item, index) => index === 2 ? {
    ...item, author: 'Grace Hopper', initials: 'GH', parent_id: items[0].id,
    anchor: { kind: 'docx', locations: [] },
  } : item);
  const content = html.replace(JSON.stringify({ schema: 'office.annotations.v1', items }),
    JSON.stringify({ schema: 'office.annotations.v1', items: replies }));
  const page = await setup(t, 1250, 1500, content);
  assert.equal(await page.locator('.office-context-card').count(), 2);
  const reply = page.locator('.office-context-card').first().locator('.office-context-reply');
  assert.equal(await reply.locator('.office-annotation-avatar').textContent(), 'GH');
  assert.equal(await reply.locator('.office-context-author').textContent(), 'Grace Hopper');
  await reply.getByRole('button', { name: 'Quote comment' }).click();
  assert.equal(await page.evaluate(() => window.annotationEvents.at(-1).item.id), items[2].id);
});

test('mobile layout keeps cards readable without floating source icons', async t => {
  const page = await setup(t, 360, 400);
  const rail = page.locator('.office-context-rail');
  assert.match(await rail.locator('summary').innerText(), /Comments.*3/);
  assert.ok((await rail.boundingBox()).width >= 350, 'collapsed mobile heading uses the available width');
  if (!await rail.evaluate(node => node.open)) await rail.locator('summary').click();
  await settle(page);
  assert.equal(await page.locator('.office-comment-marker:visible').count(), 0);
  assert.equal(await page.locator('.office-context-card:visible').count(), 3);
  assert.equal(await page.locator('.office-context-layout').evaluate(node => getComputedStyle(node).display), 'block');
});
