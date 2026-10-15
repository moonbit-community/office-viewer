import assert from 'node:assert/strict';
import { after, before, test } from 'node:test';
import { runtime } from './browser-runtime.mjs';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
let browser;
before(async () => {
  browser = await chromium.launch({ headless: true,
    ...(process.env.CHROME_EXECUTABLE ? { executablePath: process.env.CHROME_EXECUTABLE } : {}),
  });
});
after(async () => { await browser?.close(); });

async function setup(t, segmenter = true) {
  const page = await browser.newPage();
  t.after(() => page.close());
  await page.setContent('<body></body>');
  if (!segmenter) await page.evaluate(() => { Intl.Segmenter = undefined; });
  await page.addScriptTag({ content: runtime });
  return async item => page.evaluate(item => {
    const avatar = window.officeTest.avatar(item);
    document.body.replaceChildren(avatar);
    return {
      text: avatar.textContent, hidden: avatar.getAttribute('aria-hidden'),
      children: [...avatar.children].map(node => node.tagName.toLowerCase()),
      shapes: [...avatar.querySelectorAll('svg > *')].map(node => node.tagName.toLowerCase()),
      colors: [avatar.style.getPropertyValue('--office-avatar-background'), avatar.style.getPropertyValue('--office-avatar-color')],
      images: avatar.querySelectorAll('img').length,
    };
  }, item);
}

test('document initials take priority and long values stay compact', async t => {
  const avatar = await setup(t);
  assert.equal((await avatar({ author: 'Ada Lovelace', initials: ' xy ' })).text, 'XY');
  assert.equal((await avatar({ author: 'Ada Lovelace', initials: 'User' })).text, 'US');
  assert.equal((await avatar({ author: null, initials: 'AL' })).text, 'AL');
});

test('names support Latin words, Chinese names, and complete Unicode graphemes', async t => {
  const avatar = await setup(t);
  for (const [author, expected] of [
    ['Ada Lovelace', 'AL'], ['  Ada   Byron Lovelace  ', 'AL'],
    ['产品评审', '产品'], ['𠮷野', '𠮷野'], ['e\u0301mile', 'ÉM'],
  ]) assert.equal((await avatar({ author, initials: null })).text, expected);
  assert.equal((await avatar({ author: 'Engineer', initials: '👩🏽‍💻A' })).text, '👩🏽‍💻A');
});

test("an author's color stays stable across comments, replies, and initials", async t => {
  const avatar = await setup(t);
  const original = await avatar({ author: 'Ada Lovelace', initials: 'AL' });
  assert.deepEqual((await avatar({ author: 'Ada Lovelace', initials: null })).colors, original.colors);
  assert.deepEqual((await avatar({ author: ' Ada Lovelace ', initials: 'A' })).colors, original.colors);
  const backgrounds = [];
  for (const author of ['Ada Lovelace', 'Bob', '陈明', 'Lee', '产品评审']) backgrounds.push((await avatar({ author })).colors[0]);
  assert.ok(new Set(backgrounds).size > 1);
  assert.deepEqual((await avatar({ author: 'e\u0301mile', initials: null })).colors,
    (await avatar({ author: 'émile', initials: null })).colors);
});

test('missing or blank metadata produces a decorative person icon', async t => {
  const avatar = await setup(t);
  for (const item of [{}, { author: null, initials: null }, { author: ' \n ', initials: '\t' }]) {
    const result = await avatar(item);
    assert.equal(result.text, '');
    assert.equal(result.hidden, 'true');
    assert.deepEqual(result.children, ['svg']);
    assert.deepEqual(result.shapes, ['circle', 'path']);
  }
});

test('author-controlled text stays text and creates no remote image', async t => {
  const avatar = await setup(t);
  const result = await avatar({ author: '<img src=https://example.com/avatar>', initials: '<b>' });
  assert.equal(result.text, '<B');
  assert.deepEqual(result.children, []);
  assert.equal(result.hidden, 'true');
  assert.equal(result.images, 0);
});

test('browsers without grapheme segmentation preserve supplementary characters', async t => {
  const avatar = await setup(t, false);
  assert.equal((await avatar({ author: '𠮷野', initials: null })).text, '𠮷野');
  assert.deepEqual((await avatar({ author: null, initials: null })).children, ['svg']);
});
