import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { before, after, test } from 'node:test';
import { runtime } from './browser-runtime.mjs';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const styles = await readFile(new URL('../public/styles.css', import.meta.url), 'utf8');
const bytes = (await readFile(new URL('./fixtures/large-workbook.xlsx', import.meta.url))).toString('base64');
let browser;
before(async () => { browser = await chromium.launch({headless:true}); });
after(async () => { await browser?.close(); });
async function setup(t) {
  const page = await browser.newPage({viewport:{width:1200,height:900}});
  const errors=[]; page.on('pageerror', e=>errors.push(e.message));
  t.after(async()=>{await page.close(); assert.deepEqual(errors,[]);});
  await page.setContent(`<style>${styles} #document{width:900px;height:800px}</style><div id="document" class="surface sheet-surface"></div>`);
  await page.addScriptTag({content:runtime});
  await page.evaluate(bytes=>{
    window.events=[];
    window.html=officeTest.renderExcel(bytes);
    if(window.html.startsWith('ERROR:')) throw new Error(window.html);
    officeTest.install('document','excel','/docs/large.xlsx',window.html,v=>events.push(JSON.parse(v)),()=>{});
    officeTest.freeze('document');
  },bytes);
  await frame(page);
  return page;
}
async function frame(page){await page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));}
async function search(page,value){await page.evaluate(v=>officeTest.search('document',v),value);await frame(page);}
async function drag(page,start,end){
  const a=await page.locator(`td[data-cell="${start}"]`).boundingBox();
  const b=await page.locator(`td[data-cell="${end}"]`).boundingBox();
  assert.ok(a && b);
  await page.mouse.move(a.x+a.width/2,a.y+a.height/2);await page.mouse.down();
  await page.mouse.move(b.x+b.width/2,b.y+b.height/2);await page.mouse.up();
}
test('all Excel rows and columns remain reachable with bounded cell nodes',async t=>{
  const page=await setup(t);
  assert.ok(await page.locator('td[data-cell]').count()<3000);
  assert.equal(await page.locator('td[data-cell="A1"]').textContent(),'START');
  await page.locator('section[data-sheet="Sparse"] .office-sheet-viewport').evaluate(v=>{v.scrollTop=v.scrollHeight;v.scrollLeft=v.scrollWidth;});
  await frame(page);
  assert.equal(await page.locator('td[data-cell="XFD1048576"]').textContent(),'LAST_CELL');
  assert.ok(await page.locator('td[data-cell]').count()<3000);
  await page.locator('section[data-sheet="Sparse"] .office-sheet-viewport').evaluate(v=>{v.scrollTop=0;v.scrollLeft=0;});await frame(page);
  assert.equal(await page.locator('td[data-cell="A1"]').textContent(),'START');
  assert.equal(await page.locator('.truncated').count(),0);
});
test('full-data search reveals cells beyond the previous row and column limits',async t=>{
  const page=await setup(t);
  await search(page,'BEYOND_1000');
  assert.equal(await page.locator('td.office-search-hit').textContent(),'BEYOND_1000');
  await drag(page,'A1200','B1200');
  const event=await page.evaluate(()=>events.at(-1));
  assert.equal(event.range,'A1200:B1200');
  assert.equal(event.cells.find(c=>c.address==='B1200').text,'42');
  assert.equal(await page.locator('td[data-cell="B1200"]').getAttribute('data-formula'),'21*2');
  await search(page,'BEYOND_256');
  assert.equal(await page.locator('td.office-search-hit').getAttribute('data-cell'),'IW1200');
  await search(page,'LAST_CELL');
  assert.equal(await page.locator('td.office-search-hit').getAttribute('data-cell'),'XFD1048576');
});
test('merged source coordinates, frozen panes, zoom and disposal survive lazy rendering',async t=>{
  const page=await setup(t);
  await search(page,'MERGED');
  const merged=page.locator('td[data-cell="A1000"]');
  assert.equal(await merged.getAttribute('rowspan'),'3');
  assert.equal(await merged.getAttribute('colspan'),'2');
  await drag(page,'A1000','A1000');
  assert.equal(await page.evaluate(()=>events.at(-1).range),'A1000:B1002');
  await page.evaluate(()=>officeTest.zoom('document',1.5));await frame(page);
  await search(page,'LAST_CELL');
  assert.equal(await page.locator('td[data-cell="XFD1048576"]').textContent(),'LAST_CELL');
  await page.evaluate(()=>officeTest.dispose('document'));
  const count=await page.locator('td[data-cell]').count();
  await page.locator('section[data-sheet="Sparse"] .office-sheet-viewport').evaluate(v=>{v.scrollTop=0;v.scrollLeft=0;});await frame(page);
  assert.equal(await page.locator('td[data-cell]').count(),count);
});

test('range selection keeps its source anchor when scrolling recycles the original nodes',async t=>{
  const page=await setup(t);
  await search(page,'BEYOND_1000');
  const a=await page.locator('td[data-cell="A1200"]').boundingBox();
  await page.mouse.move(a.x+a.width/2,a.y+a.height/2);await page.mouse.down();
  await page.locator('section[data-sheet="Sparse"] .office-sheet-viewport').evaluate(v=>{v.scrollTop+=2000;});await frame(page);
  const target=await page.locator('td[data-cell]').evaluateAll(cells=>cells.filter(c=>+c.dataset.row>1250&&c.dataset.col==='1').map(c=>({address:c.dataset.cell,...(()=>{const r=c.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()})).find(p=>p.y>150&&p.y<500));
  assert.ok(target);
  await page.mouse.move(target.x,target.y);await page.mouse.up();
  const event=await page.evaluate(()=>events.at(-1));
  assert.equal(event.range,`A1200:${target.address}`);
  assert.equal(event.cells.find(c=>c.address==='A1200').text,'BEYOND_1000');
  await page.evaluate(()=>officeTest.clear('document'));
  await page.locator('section[data-sheet="Sparse"] .office-sheet-viewport').evaluate(v=>{v.scrollTop+=100;});await frame(page);
  assert.equal(await page.locator('.office-cell-selected').count(),0);
});

test('inactive sheets release their cell nodes and Notes navigate into an unloaded window',async t=>{
  const page=await setup(t);
  assert.equal(await page.locator('section[data-sheet="Other"] td').count(),0);
  await search(page,'SECOND_TAIL');
  assert.equal(await page.locator('td.office-search-hit').getAttribute('data-cell'),'A1500');
  assert.equal(await page.locator('section[data-sheet="Sparse"] td').count(),0);
  assert.equal(await page.getByRole("button",{name:"Show source",exact:true}).isEnabled(),true);
  const found=await page.evaluate(()=>{
    const annotations=JSON.parse(document.querySelector('.office-annotations-metadata').textContent);
    const item=annotations.items.find(item=>item.anchor.cell==='IW1200');
    return officeTest.focus('document',JSON.stringify(item));
  });
  assert.equal(found,true);await frame(page);
  assert.equal(await page.locator('td[data-cell="IW1200"] .office-note-marker').count(),1);
  await page.locator('td[data-cell="IW1200"]').click();
  assert.equal(await page.locator('[data-excel-value]').textContent(),'BEYOND_256');
  assert.equal(await page.locator('section[data-sheet="Other"] td').count(),0);
});
