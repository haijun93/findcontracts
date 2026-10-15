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
    await page.locator('#authority').fill('검증용 발주기관');
    await page.locator('#supplier').fill('검증용 공사업체');
    await page.locator('#price_unit').fill('식');
    await page.locator('#price_quantity').fill('1');
    await page.locator('#legal_basis').fill('담당자 검토용 근거');
    await page.locator('#contract_reason').fill('담당자가 검토하여 입력한 사유');
    await page.locator('#count_year').fill('2025');
    await page.locator('#department_count').fill('0');
    await page.locator('#district_count').fill('1');
    await page.locator('#count_confirmation').selectOption('확인완료');
    await page.locator('#department_signer').fill('검증용 사업부서장');
    await page.locator('#finance_signer').fill('검증용 계약부서장');
    await page.locator('#addComparison').click();
    await page.locator('#comparisons [data-field=vendor]').fill('검증용 비교업체');
    await page.locator('#comparisons [data-field=amount]').fill('3100000');
    await page.locator('#reviewed').check();
    const downloading=page.waitForEvent('download');
    await page.locator('#generate').click();
    await (await downloading).saveAs(output);
    assert.deepEqual(errors,[]);
    await page.screenshot({path:output+'.png',fullPage:true});
    console.log('PASS: real PDF upload → editable 7 sites/5 items and contract forms → reviewed values → ZIP download; no browser errors');
  } finally {await browser.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
