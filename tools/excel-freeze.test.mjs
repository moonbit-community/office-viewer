import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { after, before, test } from 'node:test';

// A module override supports CI environments that provide Playwright separately.
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

function sheet(name, { rows = 1, columns = 1, hidden = false, objectLayout = false } = {}) {
  return `<section data-sheet="${name}" data-freeze-rows="${rows}" data-freeze-columns="${columns}"
    class="${hidden ? '' : 'is-active'}" ${hidden ? 'hidden' : ''}
    style="${objectLayout ? 'min-width:1500px;min-height:1600px' : ''}">
    <table><colgroup><col style="width:44px">${'<col style="width:110px">'.repeat(12)}</colgroup>
    <thead class="office-grid-column-headings"><tr><th class="office-grid-corner"></th>
      ${Array.from({ length: 12 }, (_, col) => `<th>${String.fromCharCode(65 + col)}</th>`).join('')}</tr></thead>
    <tbody>${Array.from({ length: 40 }, (_, row) => `<tr><th class="office-row-heading">${row + 1}</th>${Array.from({ length: 12 }, (_, col) =>
      `<td data-row="${row + 1}" data-col="${col + 1}">${name}:${row + 1},${col + 1}</td>`).join('')}</tr>`).join('')}</tbody></table>
    <div class="sheet-objects"><figure data-anchor="B2">Drawing</figure></div>
    </section>`;
}

async function setup(options, metadata) {
  const page = await browser.newPage({ viewport: { width: 1000, height: 800 } });
  await page.setContent(`<style>${styles}
    #sheets { width:600px; max-width:600px; margin:20px; }
    .office-sheet-viewport { height:300px; }
    .sheet-objects { position:absolute; top:50px; left:200px; }
  </style><div id="sheets" class="surface sheet-surface">
    <div class="sheet-document office-excel-workbook">${sheet('First', options)}${sheet('Second', { rows: 2, columns: 2, hidden: true })}
      ${metadata === undefined ? '' : `<script type="application/json" class="office-freeze-metadata">${metadata}</script>`}</div>
  </div>`);
  await page.evaluate(() => {
    window.scrollCalls = 0;
    window.previousCleanupCalls = 0;
    const host = document.getElementById('sheets');
    host.__officeLayoutAnnotations = () => window.scrollCalls++;
    host.__officeViewerCleanup = () => window.previousCleanupCalls++;
  });
  await page.addScriptTag({ content: runtime });
  await page.evaluate(() => window.officeTest.freeze('sheets'));
  return page;
}

async function position(page, selector, name = 'First') {
  return page.locator(`section[data-sheet="${name}"] ${selector}`).evaluate(element => {
    const rect = element.getBoundingClientRect();
    return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
  });
}
const cell = (row, column) => `td[data-row="${row}"][data-col="${column}"]`;
function near(actual, expected, message) {
  assert.ok(Math.abs(actual - expected) < 1.5, `${message}: ${actual} vs ${expected}`);
}
async function scroll(page, name = 'First', x = 180, y = 180) {
  await page.locator(`section[data-sheet="${name}"] .office-sheet-viewport`).evaluate((viewport, point) => {
    viewport.scrollTo(point.x, point.y);
  }, { x, y });
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
}

test('freezes rows and columns independently while ordinary cells move in both directions', async t => {
  const page = await setup();
  t.after(() => page.close());
  const selectors = [cell(1, 1), cell(1, 6), cell(10, 1), cell(10, 6), 'thead th:first-child', 'tbody tr:nth-child(10) .office-row-heading'];
  const before = await Promise.all(selectors.map(selector => position(page, selector)));
  await scroll(page);
  const after = await Promise.all(selectors.map(selector => position(page, selector)));
  near(after[0].x, before[0].x, 'frozen intersection x');
  near(after[0].y, before[0].y, 'frozen intersection y');
  near(after[1].y, before[1].y, 'frozen row y');
  near(after[1].x, before[1].x - 180, 'unfrozen column x');
  near(after[2].x, before[2].x, 'frozen column x');
  near(after[2].y, before[2].y - 180, 'unfrozen row y');
  near(after[3].x, before[3].x - 180, 'ordinary cell x');
  near(after[3].y, before[3].y - 180, 'ordinary cell y');
  near(after[4].x, before[4].x, 'corner x');
  near(after[4].y, before[4].y, 'corner y');
  near(after[5].x, before[5].x, 'row number x');
  assert.equal(await page.locator('.office-sheet-grid > .sheet-objects').count(), 2);
});

test('moves preexisting object-layout section bounds into the scrolling grid', async t => {
  const page = await setup({ objectLayout: true });
  t.after(() => page.close());
  const section = page.locator('section[data-sheet="First"]');
  assert.equal(await section.evaluate(node => node.style.minWidth), '');
  assert.equal(await section.evaluate(node => node.style.minHeight), '');
  assert.ok(await section.locator('.office-sheet-viewport').evaluate(node => node.clientWidth <= 600));
  const frozen = await position(page, cell(1, 1));
  const mobile = await position(page, cell(10, 6));
  await scroll(page);
  near((await position(page, cell(1, 1))).x, frozen.x, 'frozen column after object layout');
  near((await position(page, cell(10, 6))).x, mobile.x - 180, 'ordinary column after object layout');
});

test('zoom recomputes offsets and keeps adjacent frozen rows and columns aligned', async t => {
  const page = await setup({ rows: 2, columns: 2 });
  t.after(() => page.close());
  await page.evaluate(() => {
    const host = document.getElementById('sheets');
    host.style.setProperty('--office-zoom', '1.5');
    host.__officeFreezeLayout();
  });
  const a1 = await position(page, cell(1, 1));
  const b2 = await position(page, cell(2, 2));
  await scroll(page);
  const movedA1 = await position(page, cell(1, 1));
  const movedB2 = await position(page, cell(2, 2));
  near(movedA1.x, a1.x, 'zoomed A1 x');
  near(movedA1.y, a1.y, 'zoomed A1 y');
  near(movedB2.x, b2.x, 'zoomed B2 x');
  near(movedB2.y, b2.y, 'zoomed B2 y');
  near(movedB2.x, movedA1.x + movedA1.width, 'adjacent column offsets');
  near(movedB2.y, movedA1.y + movedA1.height, 'adjacent row offsets');
});

test('switching to a previously hidden sheet lays out its own freeze settings', async t => {
  const page = await setup();
  t.after(() => page.close());
  await page.evaluate(() => {
    const sections = document.querySelectorAll('section[data-sheet]');
    sections[0].hidden = true;
    sections[0].classList.remove('is-active');
    sections[1].hidden = false;
    sections[1].classList.add('is-active');
    document.getElementById('sheets').__officeFreezeLayout();
  });
  const before = await position(page, cell(2, 2), 'Second');
  const ordinary = await position(page, cell(10, 6), 'Second');
  await scroll(page, 'Second');
  const after = await position(page, cell(2, 2), 'Second');
  const moved = await position(page, cell(10, 6), 'Second');
  near(after.x, before.x, 'second sheet B2 x');
  near(after.y, before.y, 'second sheet B2 y');
  near(moved.x, ordinary.x - 180, 'second sheet ordinary x');
  near(moved.y, ordinary.y - 180, 'second sheet ordinary y');
});

test('workbook JSON metadata takes priority over fallback data attributes', async t => {
  const page = await setup({ rows: 0, columns: 0 }, JSON.stringify({ sheets: { First: { rows: 2, columns: 2 } } }));
  t.after(() => page.close());
  const before = await position(page, cell(2, 2));
  const ordinary = await position(page, cell(3, 3));
  await scroll(page);
  const after = await position(page, cell(2, 2));
  const moved = await position(page, cell(3, 3));
  near(after.x, before.x, 'metadata freezes second column');
  near(after.y, before.y, 'metadata freezes second row');
  near(moved.x, ordinary.x - 180, 'metadata leaves third column unfrozen');
  near(moved.y, ordinary.y - 180, 'metadata leaves third row unfrozen');
});

test('malformed JSON metadata falls back to sheet attributes', async t => {
  const page = await setup({ rows: 1, columns: 1 }, '{not-json');
  t.after(() => page.close());
  const before = await position(page, cell(1, 1));
  await scroll(page);
  const after = await position(page, cell(1, 1));
  near(after.x, before.x, 'fallback frozen column');
  near(after.y, before.y, 'fallback frozen row');
});

test('resizing the host after row and column size changes refreshes sticky offsets', async t => {
  const page = await setup({ rows: 2, columns: 2 });
  t.after(() => page.close());
  await page.evaluate(() => {
    const host = document.getElementById('sheets');
    host.style.width = '550px';
    const sheet = host.querySelector('section[data-sheet="First"]');
    sheet.querySelector('colgroup col:nth-child(2)').style.width = '160px';
    sheet.querySelector('tbody tr').style.height = '60px';
  });
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const before = await position(page, cell(2, 2));
  await scroll(page);
  const a1 = await position(page, cell(1, 1));
  const b2 = await position(page, cell(2, 2));
  near(b2.x, before.x, 'resized second column');
  near(b2.y, before.y, 'resized second row');
  near(b2.x, a1.x + a1.width, 'resized adjacent columns');
  near(b2.y, a1.y + a1.height, 'resized adjacent rows');
});

test('zero or malformed freeze settings leave data cells scrollable', async t => {
  const page = await setup({ rows: '-1', columns: 'NaN' });
  t.after(() => page.close());
  const before = await position(page, cell(1, 1));
  await scroll(page);
  const after = await position(page, cell(1, 1));
  near(after.x, before.x - 180, 'invalid freeze columns ignored');
  near(after.y, before.y - 180, 'invalid freeze rows ignored');
});

test('reinstall removes old listeners and viewer cleanup stops scroll callbacks', async t => {
  const page = await setup();
  t.after(() => page.close());
  await page.evaluate(() => window.officeTest.freeze('sheets'));
  assert.equal(await page.locator('.office-sheet-viewport').count(), 2);
  await page.evaluate(() => { window.scrollCalls = 0; });
  await scroll(page);
  assert.equal(await page.evaluate(() => window.scrollCalls), 1);
  await page.evaluate(() => document.getElementById('sheets').__officeViewerCleanup());
  assert.equal(await page.evaluate(() => window.previousCleanupCalls), 1);
  assert.equal(await page.evaluate(() => document.getElementById('sheets').__officeFreezeLayout), null);
  await scroll(page, 'First', 240, 240);
  assert.equal(await page.evaluate(() => window.scrollCalls), 1);
});

test('ignores unrelated hosts and empty sheets, and gives transparent sticky cells an opaque background', async t => {
  const page = await setup();
  t.after(() => page.close());
  await page.evaluate(() => {
    const host = document.getElementById('sheets');
    const empty = document.createElement('section');
    empty.dataset.sheet = 'Empty';
    host.querySelector('.sheet-document').appendChild(empty);
    host.querySelector('td').style.backgroundColor = 'transparent';
    const unrelated = document.createElement('div');
    unrelated.id = 'unrelated';
    unrelated.innerHTML = '<section data-sheet="Unrelated"><table><tr><td>Unrelated content</td></tr></table></section>';
    document.body.appendChild(unrelated);
    const install = window.officeTest.freeze;
    install('missing');
    install('unrelated');
    install('sheets');
  });
  assert.equal(await page.locator('#unrelated .office-sheet-viewport').count(), 0);
  assert.equal(await page.locator('section[data-sheet="Empty"] .office-sheet-viewport').count(), 0);
  assert.equal(await page.locator('#sheets .office-sheet-viewport').count(), 2);
  assert.equal(await page.locator(`section[data-sheet="First"] ${cell(1, 1)}`).evaluate(element => getComputedStyle(element).backgroundColor), 'rgb(255, 255, 255)');
  const before = await position(page, cell(1, 1));
  await scroll(page);
  const after = await position(page, cell(1, 1));
  near(after.x, before.x, 'freeze survives empty sheet');
  near(after.y, before.y, 'freeze survives empty sheet');
});
