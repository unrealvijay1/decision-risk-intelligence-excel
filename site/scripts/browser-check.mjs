// Optional browser verification; uses an existing Playwright installation.
import {createRequire} from 'node:module';
import {createServer} from 'node:http';
import {readFile,stat,mkdir,writeFile} from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
const require=createRequire(import.meta.url);
const {chromium}=require(process.env.PLAYWRIGHT_PATH || 'playwright');
const root=path.resolve('site/_build');
const output=path.resolve('artifacts/site-validation');
await mkdir(output,{recursive:true});
const mime={'.html':'text/html','.css':'text/css','.js':'text/javascript','.svg':'image/svg+xml','.webp':'image/webp','.png':'image/png','.txt':'text/plain'};
const server=createServer(async(req,res)=>{
 try {
  let pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname);
  if(pathname.startsWith('/telivu/'))pathname=pathname.slice(7);
  let file=path.resolve(root,'.'+pathname);
  if(!file.startsWith(root+path.sep)&&file!==root)throw Error('Outside root');
  if((await stat(file)).isDirectory())file=path.join(file,'index.html');
  res.setHeader('Content-Type',mime[path.extname(file)]||'application/octet-stream');res.end(await readFile(file));
 }catch{res.writeHead(404);res.end('Not found');}
});
await new Promise(resolve=>server.listen(8766,'127.0.0.1',resolve));
let browser;
const results=[];
try {
 browser=await chromium.launch({headless:true,...(process.env.BROWSER_PATH?{executablePath:process.env.BROWSER_PATH}:{})});
 const page=await browser.newPage();
 const errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 const widths=[[1920,1080],[1440,900],[1366,768],[768,1024],[390,844],[375,812]];
 for(const prefix of ['/','/telivu/']){
  for(const [width,height] of widths){
   await page.setViewportSize({width,height});await page.goto('http://127.0.0.1:8766'+prefix);
   await page.evaluate(()=>document.fonts.ready);
   assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth),'Horizontal overflow at '+width);
   assert.equal(await page.locator('h1').innerText(),'Make better decisions.\nDirectly in Excel.');
   const config=await page.evaluate(()=>window.TELIVU_CONFIG);
   const enabled=config.downloadEnabled===true&&/^https:\/\//.test(config.downloadUrl);
   const expected=enabled?config.downloadUrl:'http://127.0.0.1:8766'+prefix+'index.html#download';
   const ctas=page.locator('[data-download]');assert.equal(await ctas.count(),4);
   for(const cta of await ctas.all()){
    assert.equal(await cta.innerText(),enabled?'Download Telivu Free':'Download unavailable');
    assert.equal(await cta.getAttribute('href'),expected);
   }
   if(enabled)assert.equal(await page.locator('[data-release-notes]').getAttribute('href'),config.releaseNotesUrl);
   if(width<=800){
    const toggle=page.locator('.menu-toggle');await toggle.focus();await page.keyboard.press('Enter');
    assert.equal(await toggle.getAttribute('aria-expanded'),'true');assert(await page.locator('#navigation').isVisible());
    await page.keyboard.press('Escape');assert.equal(await toggle.getAttribute('aria-expanded'),'false');assert(await toggle.evaluate(el=>el===document.activeElement));
   }
   const summary=page.locator('summary').first();await summary.focus();await page.keyboard.press('Enter');assert(await summary.evaluate(el=>el.parentElement.open));
   await page.locator('img').evaluateAll(images=>images.forEach(img=>img.loading='eager'));
   await page.waitForFunction(()=>Array.from(document.images).every(img=>img.complete&&img.naturalWidth>0));
   await page.evaluate(()=>window.scrollTo({top:0,behavior:'instant'}));
   await page.screenshot({path:path.join(output,(prefix==='/'?'root':'project')+'-'+width+'.png'),fullPage:true});
   if(prefix==='/'&&(width===1440||width===375))await page.screenshot({path:path.join(output,'hero-'+width+'.png')});
   results.push(`${prefix} ${width}×${height}: no overflow; navigation, FAQ, safe download passed`);
  }
  await page.goto('http://127.0.0.1:8766'+prefix+'resources/documentation/');
  assert((await page.locator('h1').innerText()).includes('Documentation'));
  const nestedConfig=await page.evaluate(()=>window.TELIVU_CONFIG);
  assert.equal(await page.locator('[data-download]').first().getAttribute('href'),nestedConfig.downloadEnabled?nestedConfig.downloadUrl:'http://127.0.0.1:8766'+prefix+'index.html#download');
 }
 await page.emulateMedia({reducedMotion:'reduce'});await page.goto('http://127.0.0.1:8766/');
 assert.equal(await page.locator('html').evaluate(el=>getComputedStyle(el).scrollBehavior),'auto');
 await page.route('**/site-config.js*',route=>route.fulfill({contentType:'text/javascript',body:'window.TELIVU_CONFIG={productVersion:"0.2.1",downloadEnabled:true,downloadUrl:"https://example.com/approved.exe",installerSize:"112 MB",releaseStatus:"Unsigned prerelease"};'}));
 await page.reload();assert.equal(await page.locator('[data-download]').first().innerText(),'Download Telivu Free');assert.equal(await page.locator('[data-download]').first().getAttribute('href'),'https://example.com/approved.exe');assert(await page.locator('[data-installer-size]').isVisible());
 assert((await page.locator('[data-download-status]').innerText()).includes('Unsigned prerelease'));
 await page.unroute('**/site-config.js*');
 await page.route('**/site-config.js*',route=>route.fulfill({contentType:'text/javascript',body:'window.TELIVU_CONFIG={downloadEnabled:true,downloadUrl:"file:///internal.exe"};'}));
 await page.reload();assert.equal(await page.locator('[data-download]').first().innerText(),'Download unavailable');
 await page.unroute('**/site-config.js*');
 await page.route('**/assets/product/*.webp',route=>route.abort());await page.reload();
 assert(await page.locator('.product-shot figcaption').isVisible());assert.equal(await page.locator('.placeholder').count(),0);
 // Images retain dimensions, alt text and captions if an asset cannot load.
 const noJS=await browser.newContext({javaScriptEnabled:false,viewport:{width:375,height:812}});const fallback=await noJS.newPage();await fallback.goto('http://127.0.0.1:8766/');assert(await fallback.locator('#navigation').isVisible());assert.equal(await fallback.locator('[data-download]').first().innerText(),'Download Telivu Free');await noJS.close();
 if(process.env.VERIFY_LIVE_DOWNLOAD==='1'){
  await page.unroute('**/assets/product/*.webp');
  await page.goto('http://127.0.0.1:8766/');
  const config=await page.evaluate(()=>window.TELIVU_CONFIG);
  const pending=page.waitForEvent('download',{timeout:120000});
  await page.locator('.hero [data-download]').click();
  const download=await pending;
  const installer=path.join(output,'existing-download-browser.exe');
  await download.saveAs(installer);
  assert.equal(await download.failure(),null);
  // GitHub redirects its stable release URL to a temporary release-assets URL.
  assert(['github.com','release-assets.githubusercontent.com'].includes(new URL(download.url()).hostname));
  assert.equal(download.suggestedFilename(),path.basename(new URL(config.downloadUrl).pathname));
  const bytes=await readFile(installer);
  assert.equal(bytes.length,Number(process.env.EXPECTED_INSTALLER_BYTES));
  assert.equal(createHash('sha256').update(bytes).digest('hex'),process.env.EXPECTED_INSTALLER_SHA256);
  results.push('Existing Download Flow: PASS; actual hero-button click, complete existing installer transfer and SHA-256 verified.');
 }
 assert.deepEqual(errors,[]);
 results.push('PASS: nested routes, reduced motion, enabled download, unsafe URL rejection, missing images, no-JS navigation; zero JS errors.');
 await writeFile(path.join(output,'browser-results.txt'),results.join('\n')+'\n');console.log(results.join('\n'));
}finally{if(browser)await browser.close();await new Promise(resolve=>server.close(resolve));}
