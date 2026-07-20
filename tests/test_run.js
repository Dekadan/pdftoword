// Tarayıcıda gerçek dönüştürmeyi sürer: PDF yükle -> Dönüştür -> .docx indir + paragrafları döndür
// Kullanım: NODE_PATH=$(npm root -g) node test_run.js <url> <pdf> <out.docx> <paras.json>
const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const [url, pdfPath, outDocx, parasOut] = process.argv.slice(2);
  const browser = await chromium.launch();
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();
  const errors = [];
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('pageerror', (e) => errors.push('PAGEERROR: ' + e.message));

  await page.goto(url, { waitUntil: 'load' });
  await page.waitForFunction("window.pdfjsLib && document.getElementById('go')");
  await page.setInputFiles('#file', pdfPath);
  await page.waitForFunction("!document.getElementById('go').disabled");

  const [download] = await Promise.all([
    page.waitForEvent('download'),
    page.click('#go'),
  ]);
  await download.saveAs(outDocx);
  await page.waitForFunction("window.__lastParas && window.__lastParas.length > 0", null, { timeout: 120000 });
  const paras = await page.evaluate("window.__lastParas");
  fs.writeFileSync(parasOut, JSON.stringify(paras, null, 2));

  console.log('KONSOL HATALARI:', JSON.stringify(errors));
  console.log('PARAGRAF SAYISI:', paras.length);
  await browser.close();
  if (errors.length) process.exit(2);
})().catch((e) => { console.error('DRIVER ERROR', e); process.exit(1); });
