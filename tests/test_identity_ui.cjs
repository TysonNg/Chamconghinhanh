const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname,'../static/js/app.js'),'utf8');
function section(name,next) {
  const start=source.indexOf('function '+name+'(');
  const end=source.indexOf('function '+next+'(',start+1);
  return source.slice(start,end).replace(/async\s*$/, '');
}
test('same-name employees open separate portrait collections by ID',()=>{
  const nodes={'emp-photos-modal-name':{dataset:{}}};
  let refreshed=0;
  const context=vm.createContext({
    allEmployeesList:[{employee_id:'e1',name:'Same'},{employee_id:'e2',name:'Same'}],
    activeEmpModalName:null,
    document:{getElementById:id=>nodes[id]},
    refreshEmployeePhotosModal(){refreshed++;},openModal(){}
  });
  vm.runInContext(section('openEmployeePhotosModal','refreshEmployeePhotosModal'),context);
  context.openEmployeePhotosModal('e2');
  assert.equal(context.activeEmpModalName,'e2');
  assert.equal(nodes['emp-photos-modal-name'].dataset.employeeId,'e2');
  assert.equal(refreshed,1);
});
test('portrait upload submits employee ID without asking for a reviewer',async()=>{
  const calls=[];
  const context=vm.createContext({
    activeEmpModalName:'e2',currentProjectName:'A',allProjectsList:[{name:'A',project_id:'p1'}],FormData,
    prompt:()=> { throw new Error('reviewer prompt must not open'); },showToast(){},loadPortraits:async()=>{},
    refreshEmployeePhotosModal(){},loadProjects:async()=>{},
    fetch:async(url,options)=>{calls.push({url,options});return {json:async()=>({success:true,saved_count:1})};}
  });
  vm.runInContext('async '+section('handleUploadEmployeePhoto','deleteEmployeePhoto'),context);
  await context.handleUploadEmployeePhoto({files:[new File(['x'],'face.jpg')],value:'x'});
  const form=calls[0].options.body;
  assert.equal(form.get('employee_id'),'e2');
  assert.equal(form.get('project_id'),'p1');
  assert.equal(form.get('reviewer'),'system');
  assert.equal(form.has('name'),false);
});
test('employee transfer uses today, existing payroll code, and system reviewer without prompts',async()=>{
  const calls=[];
  const context=vm.createContext({
    currentProjectName:'A',allProjectsList:[{name:'A',project_id:'p1'}],
    allEmployeesList:[{employee_id:'e2',name:'Employee',payroll_code:'001'}],
    document:{getElementById:id=>id==='transfer-employee-name'
      ? {dataset:{employeeId:'e2'}} : {value:'p2'}},
    prompt:()=> { throw new Error('transfer prompt must not open'); },
    apiPost:async(url,payload)=>{calls.push({url,payload});return {success:true};},
    showToast(){},closeModal(){},loadPortraits:async()=>{},loadProjects:async()=>{}
  });
  vm.runInContext('async '+section('submitTransferEmployee','refreshEmployeePhotosModal'),context);
  await context.submitTransferEmployee();
  assert.equal(calls[0].payload.employee_id,'e2');
  assert.equal(calls[0].payload.payroll_code,'001');
  assert.equal(calls[0].payload.reviewer,'system');
  assert.match(calls[0].payload.effective_date,/^\d{4}-\d{2}-\d{2}$/);
});
test('legacy folder mapping asks only for folder, date, and safety confirmation',async()=>{
  const calls=[];const prompts=['01','2026-09-01'];
  const context=vm.createContext({
    currentProjectName:'A',prompt:()=>prompts.shift(),confirmAction:()=>true,
    apiPost:async(url,payload)=>{calls.push({url,payload});return {success:true};},
    showToast(){},loadDailyPhotosStats(){}
  });
  vm.runInContext('async '+section('mapLegacyPhotoDay','loadDailyPhotosStats'),context);
  await context.mapLegacyPhotoDay();
  assert.equal(prompts.length,0);
  assert.equal(calls[0].payload.reviewer,'system');
  assert.equal(calls[0].payload.confirm_single_period,true);
});
test('confirmation preserves leading zeros and explicit membership date',async()=>{
  const prompts=['001','2026-01-01','Reviewer'];const calls=[];
  const context=vm.createContext({
    currentProjectName:'A',allProjectsList:[{name:'A',project_id:'p1'}],prompt:()=>prompts.shift(),showToast(){},loadPortraits:async()=>{},
    apiPost:async(url,payload)=>{calls.push({url,payload});return {success:true};}
  });
  vm.runInContext(source.slice(source.indexOf('async function confirmEmployeeIdentity(')),context);
  await context.confirmEmployeeIdentity('e2');
  assert.equal(calls[0].payload.employee_id,'e2');
  assert.equal(calls[0].payload.payroll_code,'001');
  assert.equal(calls[0].payload.valid_from,'2026-01-01');
});

test('batch internal-code button refreshes list and reports counts', async()=>{
  const calls=[];const messages=[];let refreshed=0;
  const context=vm.createContext({
    apiPost:async(url,payload)=>{calls.push({url,payload});return {success:true,created_count:3,unchanged_count:2};},
    loadPortraits:async()=>{refreshed++;},showToast:(message,type)=>messages.push({message,type})
  });
  vm.runInContext(source.slice(source.indexOf('async function generateInternalEmployeeCodes(')),context);
  const button={disabled:false};
  const pending=context.generateInternalEmployeeCodes(button);
  assert.equal(button.disabled,true);
  await context.generateInternalEmployeeCodes(button);
  await pending;
  assert.equal(calls.length,1);
  assert.equal(calls[0].url,'/api/portraits/internal-codes/generate');
  assert.equal(refreshed,1);
  assert.equal(button.disabled,false);
  assert.equal(messages[0].type,'success');
  assert.match(messages[0].message,/3/);
});
test('batch internal-code button recovers after API error',async()=>{
  const messages=[];
  const context=vm.createContext({
    apiPost:async()=>({success:false,error:'Database busy'}),
    loadPortraits:async()=>{throw new Error('Should not refresh');},
    showToast:(message,type)=>messages.push({message,type})
  });
  vm.runInContext(source.slice(source.indexOf('async function generateInternalEmployeeCodes(')),context);
  const button={disabled:false};
  await context.generateInternalEmployeeCodes(button);
  assert.equal(button.disabled,false);
  assert.equal(messages[0].message,'Database busy');
  assert.equal(messages[0].type,'error');
});
