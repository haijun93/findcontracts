'use strict';
const $ = id => document.getElementById(id);
const fields = ['title','purpose','location','scope','duration','department','budget_code','contract_method','schedule'];
const formFields=JSON.parse($('formFields').textContent);
const budgetFields = ['subtotal','adjustment','vat','total'];
const itemFields = ['name','specification','unit','quantity','unit_price','amount'];
const siteFields = ['address','note'];
const comparisonFields=['vendor','amount','note'];
let analysis = null;

function message(id, text, error=false) {
  $(id).textContent=text;
  $(id).className=error?'error':'status';
}
async function responseError(response) {
  const body=await response.json().catch(()=>({}));
  return new Error(body.error || `요청에 실패했습니다 (${response.status}).`);
}
function row(kind, value={}) {
  const keys=kind==='items'?itemFields:kind==='comparisons'?comparisonFields:siteFields;
  const tr=document.createElement('tr');
  tr.dataset.source=value.source || '사용자 입력';
  keys.forEach(key=>{
    const td=document.createElement('td');
    const input=document.createElement('input');
    input.value=value[key]??'';
    input.dataset.field=key;
    input.setAttribute('aria-label',key);
    input.maxLength=1000;
    td.append(input); tr.append(td);
  });
  const td=document.createElement('td'), remove=document.createElement('button');
  remove.textContent='삭제'; remove.type='button'; remove.className='remove';
  remove.addEventListener('click',()=>{tr.remove();$('reviewed').checked=false;});
  td.append(remove); tr.append(td); $(kind).append(tr);
}
function rows(kind) {
  return [...$(kind).children].map(tr=>Object.fromEntries([
    ...[...tr.querySelectorAll('input')].map(input=>[input.dataset.field,input.value]),
    ['source',tr.dataset.source]
  ]));
}
function showAnalysis(data) {
  analysis=data;
  for(const field of [...fields,...formFields]) $(field).value=data.facts[field]||'';
  for(const field of budgetFields) $(field).value=data.facts.budget[field]??'';
  for(const kind of ['items','sites','comparisons']) {
    $(kind).replaceChildren();
    (data.facts[kind]||[]).forEach(value=>row(kind,value));
  }
  $('warnings').replaceChildren();
  for(const warning of data.warnings) {
    const p=document.createElement('p');p.className='note';p.textContent='• '+warning;$('warnings').append(p);
  }
  $('evidence').replaceChildren();
  for(const entry of data.evidence) {
    const tr=document.createElement('tr');
    for(const field of ['field','value','source','excerpt']) {const td=document.createElement('td');td.textContent=String(entry[field]??'');tr.append(td);}
    $('evidence').append(tr);
  }
  $('sources').replaceChildren();
  for(const source of data.sources) {
    const details=document.createElement('details'), summary=document.createElement('summary'), pre=document.createElement('pre');
    summary.textContent=source.role+' · '+source.filename;
    pre.textContent=source.pages.map(p=>`[${p.page}쪽/시트]\n${p.text}`).join('\n\n');
    details.append(summary,pre);$('sources').append(details);
  }
  $('reviewed').checked=false;
  $('review').classList.remove('hidden');
}
for(const id of ['quote','request']) $(id).addEventListener('change',()=>{
  analysis=null;$('review').classList.add('hidden');$('reviewed').checked=false;
  message('message','파일을 선택했습니다. 문서 분석을 눌러 주세요.');
});
$('upload').addEventListener('submit',async event=>{
  event.preventDefault();
  const body=new FormData($('upload'));
  $('quote').disabled=true;$('request').disabled=true;
  $('analyze').disabled=true; $('review').classList.add('hidden'); analysis=null;
  message('message','문서를 분석하고 있습니다…');message('downloadMessage','');
  try {
    const response=await fetch('/api/construction/analyze',{method:'POST',body});
    if(!response.ok) throw await responseError(response);
    showAnalysis(await response.json());
    message('message','분석이 완료됐습니다. 아래 내용을 검토·수정하세요.');
  } catch(error) {message('message',error.message,true);}
  finally {$('analyze').disabled=false;$('quote').disabled=false;$('request').disabled=false;}
});
$('addComparison').addEventListener('click',()=>{row('comparisons');$('reviewed').checked=false;});
$('addItem').addEventListener('click',()=>row('items'));
$('addSite').addEventListener('click',()=>row('sites'));
$('draft').addEventListener('input',event=>{if(event.target.id!=='reviewed')$('reviewed').checked=false;});
$('draft').addEventListener('submit',async event=>{
  event.preventDefault();if(!analysis)return;
  const facts=Object.fromEntries([...fields,...formFields].map(field=>[field,$(field).value]));
  facts.budget=Object.fromEntries(budgetFields.map(field=>[field,$(field).value]));
  facts.items=rows('items');facts.sites=rows('sites');facts.comparisons=rows('comparisons');
  $('generate').disabled=true; $('analyze').disabled=true;
  $('quote').disabled=true;$('request').disabled=true;
  message('downloadMessage','검토값으로 초안을 작성하고 있습니다…');
  try {
    const response=await fetch('/api/construction/generate',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({facts,reviewed:$('reviewed').checked,provenance:{evidence:analysis.evidence,sources:analysis.sources}})});
    if(!response.ok)throw await responseError(response);
    const blob=await response.blob(),url=URL.createObjectURL(blob),anchor=document.createElement('a');
    anchor.href=url;anchor.download='공사문서_초안.zip';document.body.append(anchor);anchor.click();anchor.remove();
    setTimeout(()=>URL.revokeObjectURL(url),60000);
    message('downloadMessage','초안 ZIP을 작성했습니다. 다운로드 폴더에서 Word 문서와 검토기록을 확인하세요.');
  } catch(error){message('downloadMessage',error.message,true);}
  finally {$('generate').disabled=false;$('analyze').disabled=false;$('quote').disabled=false;$('request').disabled=false;}
});
