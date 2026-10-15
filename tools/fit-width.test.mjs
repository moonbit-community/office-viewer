import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { before, after, test } from 'node:test';
import { runtime } from './browser-runtime.mjs';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const styles = await readFile(new URL('../public/styles.css', import.meta.url), 'utf8');
let browser;
before(async () => { browser = await chromium.launch({headless:true, ...(process.env.CHROME_EXECUTABLE ? {executablePath:process.env.CHROME_EXECUTABLE} : {})}); });
after(async () => { await browser?.close(); });
const frame = page => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
async function setup(t) {
  const page = await browser.newPage({viewport:{width:1200,height:900}});
  t.after(() => page.close());
  await page.setContent(`<style>${styles} #document {width:500px;height:700px;overflow:auto;padding:12px;box-sizing:border-box}</style><div id="document" class="office-viewer-surface"></div>`);
  await page.addScriptTag({content:runtime});
  await page.evaluate(() => {
    window.zoomEvents=[];
    officeTest.install('document','word','/review.docx','<div class="word-pages"><div class="word-page"><svg width="800" height="1100"><text x="40" y="80">Selectable document text</text></svg></div></div>',()=>{},()=>{});
    officeTest.fit('document','word',true, zoom => zoomEvents.push(zoom));
  });
  await frame(page);
  return page;
}
async function geometry(page) {
  return page.evaluate(() => {
    const host=document.getElementById('document');
    const paper=host.querySelector('.word-page');
    return {zoom:zoomEvents.at(-1),pageWidth:paper.getBoundingClientRect().width,available:host.clientWidth-24};
  });
}
test('fit width keeps the complete page visible and follows host resizing', async t => {
  const page=await setup(t);
  let size=await geometry(page);
  assert.ok(size.pageWidth<=size.available && size.zoom<100);
  await page.evaluate(() => {document.getElementById('document').style.width='320px';});
  await frame(page);
  const narrow=await geometry(page);
  assert.ok(narrow.zoom<size.zoom && narrow.pageWidth<=narrow.available);
  await page.evaluate(() => {document.getElementById('document').style.width='1100px';});
  await frame(page);
  assert.equal((await geometry(page)).zoom,100,'fit does not enlarge small pages beyond actual size');
});
test('manual zoom survives resizing, and re-enabling fit resumes adaptation', async t => {
  const page=await setup(t);
  await page.evaluate(() => {officeTest.fit('document','word',false,()=>{});officeTest.zoom('document',0.75);document.getElementById('document').style.width='320px';});
  await frame(page);
  assert.equal(await page.locator('.word-page').evaluate(el => el.getBoundingClientRect().width),600);
  await page.evaluate(() => officeTest.fit('document','word',true,zoom=>zoomEvents.push(zoom)));
  await frame(page);
  const fitted=await geometry(page);
  assert.ok(fitted.pageWidth<=fitted.available);
});
test('disposal retires width observers and pending callbacks', async t => {
  const page=await setup(t);
  const count=await page.evaluate(() => {officeTest.dispose('document');document.getElementById('document').style.width='300px';return zoomEvents.length;});
  await frame(page);
  assert.equal(await page.evaluate(()=>zoomEvents.length),count);
});
test('pinch changes zoom while ordinary wheel scrolling preserves it', async t => {
  const page=await setup(t);
  await page.evaluate(() => {
    officeTest.fit('document','word',false,()=>{});
    officeTest.pinch('document','word',10,200,zoom=>zoomEvents.push(zoom));
    document.getElementById('document').dispatchEvent(new WheelEvent('wheel',{deltaY:-10,ctrlKey:true,cancelable:true,clientX:100,clientY:200}));
  });
  await frame(page);
  const pinched=await page.evaluate(()=>zoomEvents.at(-1));
  await page.evaluate(()=>document.getElementById('document').dispatchEvent(new WheelEvent('wheel',{deltaY:50,cancelable:true})));
  await frame(page);
  assert.equal(await page.evaluate(()=>zoomEvents.at(-1)),pinched);
  assert.ok(pinched>59);
});
