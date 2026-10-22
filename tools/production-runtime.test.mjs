import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { extname } from 'node:path';
import { test } from 'node:test';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');

// Exercise Warren's shipped entry as well as its separately loaded runtime.
// An unscoped compiled MoonBit `Set` otherwise shadows the browser's native Set.
test('production PPTX entry retains native globals, renders slides and navigates thumbnails', async t => {
  const browser = await chromium.launch({headless:true});
  t.after(() => browser.close());
  const page = await browser.newPage({viewport:{width:1400,height:900}});
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
  await page.route('http://production.test/**', async route => {
    const path = new URL(route.request().url()).pathname;
    try {
      const body = await readFile(new URL(`../dist${path === '/' ? '/index.html' : path}`, import.meta.url));
      const contentType = {'.js':'text/javascript','.css':'text/css','.html':'text/html','.ttf':'font/ttf','.woff2':'font/woff2','.wasm':'application/wasm','.svg':'image/svg+xml'}[extname(path)] || (path === '/' ? 'text/html' : 'application/octet-stream');
      await route.fulfill({body,contentType});
    } catch {
      await route.fulfill({status:404,body:'Not found'});
    }
  });
  await page.goto('http://production.test/');
  await page.getByRole('button',{name:'Open file',exact:true}).waitFor();
  assert.equal(await page.evaluate(() => typeof new Set().clear), 'function');
  await page.locator('#file-input').setInputFiles(new URL('./fixtures/pptx-preview.pptx', import.meta.url).pathname);
  await page.locator('.pptx-slide').nth(1).waitFor();
  await page.waitForFunction(() => document.querySelectorAll('#presentation-slide-list button').length === 2);
  await page.locator('#presentation-slide-list button').nth(1).click();
  await page.waitForFunction(() => document.querySelectorAll('#presentation-slide-list button')[1].getAttribute('aria-current') === 'page');
  await page.getByRole('searchbox').fill('Second');
  await page.locator('.pptx-slide text.office-search-hit').waitFor();
  assert.equal(await page.locator('.pptx-slide').count(), 2);
  assert.equal(await page.locator('.pptx-slide image').count(), 2);
  assert.equal(await page.locator('.office-font-substitutions').count(), 1);
  assert.match(await page.locator('.office-font-substitutions').textContent(), /Arial/);
  assert.match(await page.locator('.office-font-substitutions summary').textContent(), /^Font substitutions \(\d+\)$/);
  assert.deepEqual(errors, []);
});
