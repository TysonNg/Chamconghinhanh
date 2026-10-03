const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const read=name=>fs.readFileSync(path.join(__dirname,'../static/js',name),'utf8');
test('all native confirmations are replaced and no times are synthesized',()=>{
 for(const name of ['app.js','supplement-batches.js']) assert.doesNotMatch(read(name), /\bconfirm\(/);
 assert.doesNotMatch(read('supplement-batches.js'), /getRandomShiftTime|target_time: timeStr/);
});
test('HTTP errors cannot report success through apiPost',async()=>{
 const source=read('app.js'); const start=source.indexOf('async function parseJsonResponse'); const end=source.indexOf('// ==================== Vietnamese',start);
 const ctx=vm.createContext({fetch:async()=>({ok:false,status:500,text:async()=>JSON.stringify({success:true})})});
 vm.runInContext(source.slice(start,end),ctx); assert.equal((await ctx.apiPost('/delete')).success,false);
});
test('shared confirmation module exists in both pages',()=>{
 assert.ok(fs.existsSync(path.join(__dirname,'../static/js/confirm-action.js')));
 for(const name of ['index.html','supplement_batches.html']) assert.match(fs.readFileSync(path.join(__dirname,'../templates',name),'utf8'),/confirm-action\.js/);
});

test('failed generation status never produces a success notification',async()=>{
 let source=read('supplement-batches.js');source=source.slice(0,source.indexOf('  // --- INITIALIZE ---'))+'\n window.regenerate=regenerateStagingPhoto; })();';
 const message={};const ctx=vm.createContext({window:{},document:{getElementById:id=>id==='batch-message'?message:null},console,fetch:async()=>({ok:true,json:async()=>({success:true,processing:{watermark:{status:'failed'}},status:'verification_failed'})})});
 vm.runInContext(source,ctx);assert.equal(await ctx.window.regenerate('one',{innerHTML:'Generate'}),false);assert.equal(message.className.includes('success'),false);
});

test('historical employee loading uses project ID and selects employee ID',async()=>{
 const calls=[];const hidden={value:''};let source=read('supplement-batches.js');source=source.slice(0,source.indexOf('  // --- INITIALIZE ---'))+'\n currentProject="p1"; window.loadEmployees=loadEmployees; })();';
 const ctx=vm.createContext({window:{},document:{getElementById:id=>id==='supp-employee-val'?hidden:null},console,fetch:async url=>{calls.push(url);return {ok:true,json:async()=>({success:true,employees:[{employee_id:'historical-e2',name:'Same name'}]})};}});
 vm.runInContext(source,ctx);await ctx.window.loadEmployees('p1');assert.equal(calls[0],'/api/portraits?include_history=true&project_id=p1');assert.equal(hidden.value,'historical-e2');
});

function recordHarness(item, accepted=true) {
 const nodes={'batch-message':{},'supp-project-select':{value:'p1'},'supp-employee-val':{value:'e2'}};const calls=[];const confirmations=[];
 let source=read('supplement-batches.js');source=source.slice(0,source.indexOf('  // --- INITIALIZE ---'))+'\n currentProject="p1"; projectsList=[{project_id:"p1",name:"Project One"}]; currentSelectedEmployee={employee_id:"e2",name:"Employee Two"}; window.bind=rebindStagingPhoto; window.check=recheckStagingPhoto; })();';
 const ctx=vm.createContext({window:{},document:{getElementById:id=>nodes[id]||null},console,confirmAction:async text=>{confirmations.push(text);return accepted;},fetch:async(url,options)=>{calls.push({url,options});return {ok:true,json:async()=>({success:true,item:{...item,can_apply:true}})};}});
 vm.runInContext(source,ctx);return {ctx,nodes,calls,confirmations};
}
test('legacy rebinding confirms names and explicitly discloses source baseline',async()=>{
 const item={id:'legacy',source_sha256:null};const h=recordHarness(item);const button={disabled:false};await h.ctx.window.bind(item,button);
 assert.equal(h.calls[0].url,'/api/supplement/records/legacy');assert.equal(h.calls[0].options.method,'PATCH');
 assert.deepEqual(JSON.parse(h.calls[0].options.body),{project_id:'p1',employee_id:'e2',confirm_legacy_source:true});
 assert.match(h.confirmations[0],/Project One/);assert.match(h.confirmations[0],/Employee Two/);assert.match(h.confirmations[0],/baseline|m?c|ban ??u/i);assert.equal(button.disabled,false);
});
test('hashed records omit baseline flag and all absent time fields',async()=>{
 const item={id:'hashed',source_sha256:'abc'};const h=recordHarness(item);await h.ctx.window.bind(item,{disabled:false});
 assert.deepEqual(JSON.parse(h.calls[0].options.body),{project_id:'p1',employee_id:'e2'});
});
test('cancelled rebinding and applied records send no mutations',async()=>{
 for(const [item,accepted] of [[{id:'cancel'},false],[{id:'applied',applied_path:'committed.jpg'},true]]){
  const h=recordHarness(item,accepted);await h.ctx.window.bind(item,{disabled:false});assert.equal(h.calls.length,0);
 }
});
test('recheck calls dedicated evidence endpoint and restores button',async()=>{
 const item={id:'one'};const h=recordHarness(item);const button={disabled:false};await h.ctx.window.check(item,button);
 assert.equal(h.calls[0].url,'/api/supplement/records/one/check');assert.equal(h.calls[0].options.method,'POST');assert.equal(button.disabled,false);
});

test('recheck deduplicates concurrent requests and recovers from HTTP failure',async()=>{
 const h=recordHarness({id:'one'});const button={disabled:false};
 h.ctx.fetch=async(url,options)=>{h.calls.push({url,options});return {ok:false,status:422,json:async()=>({success:true,error:'conditions failed'})};};
 await Promise.all([h.ctx.window.check({id:'one'},button),h.ctx.window.check({id:'one'},button)]);
 assert.equal(h.calls.length,1);assert.equal(button.disabled,false);assert.equal(h.nodes['batch-message'].className.includes('error'),true);
});
