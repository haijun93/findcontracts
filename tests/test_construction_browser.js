// Optional real-browser integration check; sample files remain outside the repository.
// node tests/test_construction_browser.js BASE_URL QUOTE_PDF REQUEST_PDF OUTPUT_ZIP
const { chromium } = require('playwright');
const assert = require('node:assert/strict');
(async () => {
  const [base, quote, request, output] = process.argv.slice(2);
  if(!output) throw new Error('Provide BASE_URL QUOTE_PDF REQUEST_PDF OUTPUT_ZIP');
  const browser = await chromium.launch({executablePath: process.env.CHROMIUM_PATH || '/usr/bin/chromium', headless: true, args:['--no-sandbox']});
  try {
    const page = await browser.newPage({acceptDownloads:true,viewport:{width:1400,height:1000}});
    const errors=[];page.on('pageerror',error=>errors.push(error.message));
    await page.goto(base+'/construction');
    await page.locator('#quote').setInputFiles(quote);
    await page.locator('#request').setInputFiles(request);
    await page.locator('#analyze').click();
    await page.locator('#review:not(.hidden)').waitFor();
    assert.equal(await page.locator('#total').inputValue(),'2970000');
    assert.equal(await page.locator('#sites tr').count(),7);
    assert.equal(await page.locator('#items tr').count(),5);
    assert.equal(await page.locator('#duration').inputValue(),'');
    await page.locator('#purpose').fill('의뢰 내용에 따른 지주식 안내판 철거 및 전기 마감');
    await page.locator('#duration').fill('착공일로부터 12일');
    await page.locator('#reviewed').check();
    await page.locator('#duration').fill('착공일로부터 14일');
    assert.equal(await page.locator('#reviewed').isChecked(),false);
    await page.locator('#reviewed').check();
    const downloading=page.waitForEvent('download');
    await page.locator('#generate').click();
    await (await downloading).saveAs(output);
    assert.deepEqual(errors,[]);
    await page.screenshot({path:output+'.png',fullPage:true});
    console.log('PASS: real PDF upload → editable 7 sites/5 items → reviewed period → ZIP download; no browser errors');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
