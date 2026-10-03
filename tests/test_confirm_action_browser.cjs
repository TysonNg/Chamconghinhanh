const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const puppeteer = require('../zalo-service/node_modules/puppeteer-core');
const script = name => fs.readFileSync(path.join(__dirname,'../static/js',name),'utf8');
test('real Chrome confirmation lifecycle and mocked staging deletion', {timeout:30000}, async () => {
 const browser = await puppeteer.launch({executablePath:'C:/Program Files/Google/Chrome/Application/chrome.exe',headless:true,args:['--no-sandbox']});
 try {
  const page=await browser.newPage();
  await page.setContent('<button id="trigger">Delete</button><button id="outside">Outside</button><div id="batch-message"></div><button id="btn-delete-all-staging">Delete all</button>');
  await page.addScriptTag({content:script('confirm-action.js')});
  const open=async()=>{await page.evaluate(()=>{document.querySelector('#trigger').focus();document.body.style.overflow='auto';window.result=null;confirmAction('<img src=x> Delete?').then(v=>window.result=v);});};
  for(const mode of ['cancel','escape','outside','accept']) {
   await open();
   assert.equal(await page.$eval('#confirm-action-message',el=>el.textContent),'<img src=x> Delete?');
   assert.equal(await page.$eval('#confirm-action-message',el=>el.children.length),0);
   assert.equal(await page.evaluate(()=>confirmAction('duplicate')),false);
   assert.equal(await page.$$eval('[role=alertdialog]',els=>els.length),1);
   await page.focus('#outside');
   assert.equal(await page.evaluate(()=>document.activeElement.dataset.confirm),'cancel');
   await page.keyboard.press('Tab');
   assert.equal(await page.evaluate(()=>document.activeElement.dataset.confirm),'accept');
   if(mode==='escape') await page.keyboard.press('Escape');
   else if(mode==='outside') await page.click('#confirm-action-backdrop',{offset:{x:2,y:2}});
   else await page.click(`[data-confirm="${mode}"]`);
   assert.equal(await page.evaluate(()=>window.result),mode==='accept');
   assert.equal(await page.$('#confirm-action-backdrop'),null);
   assert.equal(await page.evaluate(()=>document.body.style.overflow),'auto');
   assert.equal(await page.evaluate(()=>document.activeElement.id),'trigger');
  }
  // Exercise an app.js delete handler with its real shared request guard.
  const app=script('app.js');
  const api=app.slice(app.indexOf('async function parseJsonResponse'),app.indexOf('// ==================== Vietnamese'));
  const deletion=app.slice(app.indexOf('async function deleteResultFile'),app.indexOf('\n//',app.indexOf('async function deleteResultFile')));
  const guardStart=app.indexOf('// Keep destructive actions pending');
  const guards=app.slice(guardStart,app.indexOf('})();',guardStart)+5);
  const guardedNames=[...guards.matchAll(/(\w+) = guard\(/g)].map(match=>match[1]);
  await page.addScriptTag({content:guardedNames.filter(name=>name!=='deleteResultFile').map(name=>`function ${name}(){}`).join('\n')});
  await page.addScriptTag({content:api+'\n'+deletion+'\n'+guards});
  await page.evaluate(()=>{window.appCalls=0;window.showToast=()=>{};window.loadResults=()=>{};window.fetch=async()=>{appCalls++;return {ok:false,status:500,text:async()=>JSON.stringify({success:true,error:'mock failure'})};}; document.querySelector('#trigger').onclick=()=>deleteResultFile('test.docx');});
  await page.click('#trigger');
  await page.evaluate(()=>deleteResultFile('test.docx'));
  assert.equal(await page.$eval('#trigger',b=>b.disabled),true);
  await page.click('[data-confirm=accept]');
  await page.waitForFunction(()=>!document.querySelector('#trigger').disabled);
  assert.equal(await page.evaluate(()=>appCalls),1);
  assert.equal(await page.evaluate(()=>document.activeElement.id),'trigger');
  let source=script('supplement-batches.js');
  source=source.slice(0,source.indexOf('  // --- INITIALIZE ---'))+`\n stagingPhotos=[{id:'visible-1'},{id:'visible-2'}]; setupStagingActions(); })();`;
  await page.evaluate(()=>{window.calls=[];window.responseOK=false;window.responseSuccess=true;window.fetch=async(url,options)=>{calls.push({url,body:JSON.parse(options.body)});return {ok:responseOK,json:async()=>({success:responseSuccess,error:'mock failure'})};};});
  await page.addScriptTag({content:source});
  await page.click('#btn-delete-all-staging');
  await page.evaluate(()=>document.querySelector('#btn-delete-all-staging').onclick());
  await page.click('[data-confirm=cancel]');
  assert.equal(await page.evaluate(()=>calls.length),0);
  assert.equal(await page.$eval('#btn-delete-all-staging',b=>b.disabled),false);
  for(const [ok,success] of [[false,true],[true,false],[true,true]]) {
   await page.evaluate(([ok,success])=>{responseOK=ok;responseSuccess=success;},[ok,success]);
   await page.click('#btn-delete-all-staging');
   await page.click('[data-confirm=accept]');
   await page.waitForFunction(()=>!document.querySelector('#btn-delete-all-staging').disabled);
   assert.deepEqual(await page.evaluate(()=>calls.at(-1).body),{ids:['visible-1','visible-2']});
   assert.equal(await page.$eval('#batch-message',el=>el.classList.contains('error')),!(ok&&success));
  }
  assert.equal(await page.evaluate(()=>calls.length),3);
  // Render a real card, test all separate statuses, and exercise single deletion.
  await page.evaluate(()=>{const grid=document.createElement('div');grid.id='supp-staging-grid';document.body.append(grid);});
  let cards=script('supplement-batches.js');
  cards=cards.slice(0,cards.indexOf('  // --- INITIALIZE ---'))+'\n currentProject="p1"; projectsList=[{project_id:"p1",name:"Project One"}]; currentSelectedEmployee={employee_id:"e2",name:"Employee Two"}; window.loadCards=loadStaging; })();';
  await page.evaluate(()=>{window.fetch=async(url,options)=>{
    if(options) { calls.push({url,body:JSON.parse(options.body)}); return {ok:false,json:async()=>({success:true,error:'single failure'})}; }
    return {ok:true,json:async()=>({success:true,items:[{id:'one',project_id:'p1',employee_id:'e1',original_name:'one.jpg',url:'data:,',target_date:'2026-10-04',date_status:'consistent',face_status:'matched',integrity_status:'intact',storage_status:'available',processing:{watermark:'not_requested',exif:'not_requested'},can_apply:false,reasons:['Blocked reason']}]})};
  };});
  await page.addScriptTag({content:cards}); await page.evaluate(()=>loadCards());
  const text=await page.$eval('#supp-staging-grid',el=>el.textContent);
  for(const value of ['consistent','matched','intact','available','not_requested','Blocked reason']) assert.ok(text.includes(value));
  assert.equal(await page.$eval('.supp-action-main-btn',b=>b.disabled),true);
  await page.click('.supp-staging-card-actions button.btn-danger');
  await page.click('[data-confirm=accept]');
  await page.waitForFunction(()=>!document.querySelector('.supp-staging-card-actions button.btn-danger').disabled);
  assert.deepEqual(await page.evaluate(()=>calls.at(-1).body),{ids:['one']});
  assert.equal(await page.$eval('#batch-message',el=>el.classList.contains('error')),true);
  await page.evaluate(()=>{
    const project=document.createElement('select');project.id='supp-project-select';project.innerHTML='<option value="p1">Project One</option>';document.body.append(project);
    const employee=document.createElement('input');employee.id='supp-employee-val';employee.value='e2';document.body.append(employee);
    window.cardItem={id:'legacy',project_id:'p1',employee_id:'e1',source_sha256:null,url:null,original_name:'source.jpg',date_status:'unknown',face_status:'not_run',integrity_status:'unverified',storage_status:'missing',can_apply:false,reasons:['Check required'],processing:{watermark:'not_requested',exif:'not_requested'}};
    window.fetch=async(url,options)=>{
      if(options) {
        calls.push({url,method:options.method,body:JSON.parse(options.body)});
        cardItem.can_apply=url.endsWith('/check');
        if(options.method==='PATCH') cardItem.source_sha256='baseline-hash';
        return {ok:true,json:async()=>({success:true,item:cardItem})};
      }
      return {ok:true,json:async()=>({success:true,items:[cardItem]})};
    };
  });
  await page.evaluate(()=>loadCards());
  assert.equal(await page.$('.supp-staging-thumb-img'),null);
  assert.equal(await page.$eval('a[aria-disabled]',el=>el.hasAttribute('href')),false);
  assert.ok((await page.$eval('.supp-staging-missing-image',el=>el.textContent)).length>0);
  assert.equal(await page.$eval('.supp-action-main-btn',b=>b.disabled),true);
  await page.click('.btn-recheck-evidence');
  await page.waitForFunction(()=>!document.querySelector('.supp-action-main-btn').disabled);
  assert.equal(await page.evaluate(()=>calls.at(-1).url),'/api/supplement/records/legacy/check');
  await page.click('.btn-rebind-identity');
  const confirmation=await page.$eval('#confirm-action-message',el=>el.textContent);
  for(const text of ['Project One','Employee Two','baseline']) assert.ok(confirmation.includes(text));
  await page.click('[data-confirm=accept]');
  await page.waitForFunction(()=>document.querySelector('.supp-action-main-btn').disabled);
  assert.deepEqual(await page.evaluate(()=>calls.at(-1).body),{project_id:'p1',employee_id:'e2',confirm_legacy_source:true});
  await page.evaluate(()=>{cardItem.applied_path='immutable.jpg';return loadCards();});
  assert.equal(await page.$eval('.btn-rebind-identity',b=>b.disabled),true);


 } finally { await browser.close(); }
});
