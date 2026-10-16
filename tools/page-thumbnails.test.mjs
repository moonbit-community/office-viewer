import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { before, after, test } from 'node:test';
import { runtime } from './browser-runtime.mjs';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const module = await readFile(new URL('../public/page-thumbnails.js', import.meta.url), 'utf8');
const css = await readFile(new URL('../public/page-thumbnails.css', import.meta.url), 'utf8');
let browser;
before(async () => { browser = await chromium.launch({headless:true, ...(process.env.CHROME_EXECUTABLE ? {executablePath:process.env.CHROME_EXECUTABLE} : {})}); });
after(async () => { await browser?.close(); });
const frame = page => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const word = `<div class="word-pages">${Array.from({length:6},(_,i)=>`<div class="word-page" style="height:600px"><svg width="400" height="600" viewBox="0 0 400 600"><defs><clipPath id="clip-${i}"><rect width="400" height="600"/></clipPath></defs><g clip-path="url(#clip-${i})"><text x="30" y="60" font-family="serif">Chapter ${i+1}</text></g></svg></div>`).join('')}</div>`;
async function setup(t, kind='word') {
  const page = await browser.newPage({viewport:{width:1000,height:800}});
  t.after(() => page.close());
  await page.route('http://viewer.test/**', route => {
    const path = new URL(route.request().url()).pathname;
    route.fulfill({contentType:path.endsWith('.js')?'text/javascript':'text/html',body:path.endsWith('.js')?module:`<style>${css} .layout{display:flex;height:700px} #gallery{width:180px;flex:none} #document{width:600px;height:700px;overflow:auto} .word-page{background:white;margin:20px}</style><div class="layout"><div id="gallery" class="office-page-list" data-page-label="Chapter {page}" data-unavailable-label="Unavailable"></div><div id="document"></div></div>`});
  });
  await page.goto('http://viewer.test/');
  await page.addScriptTag({content:runtime});
  await page.evaluate(async ({kind,word}) => {
    window.pageModule = await import('/page-thumbnails.js');
    window.mountPages = enabled => pageModule.mount(document.getElementById('document'),document.getElementById('gallery'),{kind,key:'/demo',enabled,systemFonts:false});
    if (kind === 'word') officeTest.install('document','word','/demo.docx',word,()=>{},()=>{});
    window.disposed = () => officeTest.dispose('document');
  },{kind,word});
  return page;
}
async function painted(page, number) {
  await page.waitForFunction(number=>document.querySelector(`[data-page="${number}"]`)?.dataset.painted==='true',number);
}

test('a standalone host gets isolated Word previews, custom labels and document navigation',async t=>{
  const page = await setup(t);
  await page.evaluate(()=>mountPages(true));
  await painted(page,1);
  assert.equal(await page.getByRole('button',{name:'Chapter 1',exact:true}).count(),1);
  const isolation=await page.evaluate(()=>{
    const source=document.querySelector('#document svg'),copy=document.querySelector('#gallery svg');
    return {source:source.querySelector('[id]').id,copy:copy.querySelector('[id]').id,reference:copy.querySelector('g').getAttribute('clip-path'),family:copy.querySelector('text').style.fontFamily};
  });
  assert.notEqual(isolation.copy,isolation.source);
  assert.equal(isolation.reference,`url(#${isolation.copy})`);
  assert.equal(isolation.family,'serif');
  await page.getByRole('button',{name:'Chapter 4',exact:true}).click();
  await page.waitForFunction(()=>document.querySelector('[data-page="4"]').getAttribute('aria-current')==='page');
  assert.ok(await page.locator('#document').evaluate(node=>node.scrollTop)>1500);
  await page.getByRole('button',{name:'Chapter 4',exact:true}).press('Home');
  await page.waitForFunction(()=>document.querySelector('[data-page="1"]').getAttribute('aria-current')==='page');
  await page.evaluate(()=>{document.getElementById('gallery').dataset.pageLabel='第 {page} 页';document.getElementById('document').scrollTop=1220;});
  await page.getByRole('button',{name:'第 3 页',exact:true}).waitFor();
  await page.waitForFunction(()=>document.querySelector('[data-page="3"]').getAttribute('aria-current')==='page');
});

test('Word document replacement retires thumbnails through the normal surface disposal hook',async t=>{
  const page = await setup(t);
  await page.evaluate(()=>{
    const host=document.getElementById('document');
    host.__officeViewerPageThumbnailsCleanup=()=>host.__officeViewerPageThumbnails?.dispose();
    mountPages(true);
  });
  await painted(page,1);
  await page.evaluate(()=>officeTest.install('document','word','/next.docx','<div class="word-pages"><div class="word-page"><svg width="400" height="600"><text>Next document</text></svg></div></div>',()=>{},()=>{}));
  assert.equal(await page.locator('#gallery button').count(),0);
  await page.evaluate(()=>mountPages(true));
  await painted(page,1);
  assert.equal(await page.locator('#gallery button').count(),1);
  assert.equal(await page.locator('#gallery svg').textContent(),'Next document');
  await page.evaluate(()=>disposed());
  assert.equal(await page.locator('#gallery button').count(),0);
});

test('retired PDF page requests cannot paint into a replacement list',async t=>{
  const page = await setup(t,'pdf');
  await page.evaluate(()=>{
    window.renderCalls=0;
    const host=document.getElementById('document');
    host.innerHTML='<div class="pdf-container"></div>';
    window.pending=[];
    host.__officeViewerPdf={alive:true,doc:{numPages:3,getPage:()=>new Promise(resolve=>pending.push(resolve))},viewer:{currentPageNumber:1},bus:{on(){},off(){}}};
    mountPages(true);
  });
  await page.waitForFunction(()=>pending.length===2);
  await page.evaluate(()=>{
    document.getElementById('document').__officeViewerPageThumbnails.dispose();
    document.getElementById('gallery').textContent='Replacement content';
    for(const resolve of pending) resolve({render(){renderCalls++;throw Error('retired page rendered');}});
  });
  await frame(page);
  assert.equal(await page.evaluate(()=>renderCalls),0);
  assert.equal(await page.locator('#gallery').textContent(),'Replacement content');
});

test('a collapsed thumbnail list does not render pages until its host enables it',async t=>{
  const page = await setup(t);
  await page.evaluate(()=>mountPages(false));
  await frame(page);
  assert.equal(await page.locator('#gallery button').count(),0);
  await page.evaluate(()=>mountPages(true));
  await painted(page,1);
  assert.equal(await page.locator('#gallery button').count(),6);
});
