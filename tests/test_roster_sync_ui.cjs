const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync('static/js/app.js', 'utf8');
const start = source.indexOf('async function checkRosterBeforeScan(');
const end = source.indexOf('async function syncUploadedRoster(', start);
async function run(transfer, approved = false) {
    let review = false;
    const requests = [];
    const ctx = {
        allProjectsList: [{name: 'Target', project_id: 'target-id'}],
        apiGet: async () => ({files: [{name: 'month.xlsx'}]}),
        apiPost: async (url, body) => {
            requests.push({url, body});
            return {success: true, counts: {transfer}, sync_approved: approved};
        },
        openEmployeeAudit: async () => {},
        reviewRosterTransfers: async () => {review = true;},
        syncUploadedRoster: async () => {review = true;},
        previewUploadedRoster: async () => {},
        document: {getElementById: () => ({})},
        showToast: () => {},
    };
    vm.createContext(ctx);
    vm.runInContext(source.slice(start, end), ctx);
    const proceed = await ctx.checkRosterBeforeScan('month', 'excel', 'Target');
    assert.equal(proceed, true);
    assert.equal(review, false);
    assert.equal(requests[0].body.filename, 'month.xlsx');
    assert.equal(requests[0].body.project_id, 'target-id');
}
(async () => {
    await run(1);
    await run(0);
    await run(1, true);
    assert.match(source, /syncUploadedRoster\('\$\{safeName\}', 'excel'\)/);
    assert.match(source, /syncUploadedRoster\('\$\{safeName\}', 'pdf'\)/);
    const syncStart = source.indexOf('async function syncUploadedRoster(');
    const syncEnd = source.indexOf('async function previewUploadedRoster(', syncStart);
    const elements = {};
    const ctx = {
        allProjectsList: [{name: 'A', project_id: 'a'}],
        document: {getElementById: id => elements[id] ||= {
            replaceChildren(option) {this.options = [option];},
            add(option) {this.options.push(option);},
        }},
        Option: function(label, value) {this.label = label; this.value = value;},
        openModal: id => {ctx.modal = id;},
        uploadedRosterSync: null,
    };
    vm.createContext(ctx);
    vm.runInContext(source.slice(syncStart, syncEnd), ctx);
    await ctx.syncUploadedRoster('month.xlsx', 'excel');
    assert.equal(ctx.modal, 'modal-roster-sync');
    assert.equal(elements['roster-sync-project'].options[0].value, '');
    assert.equal(elements['roster-sync-apply'].disabled, true);
    assert.equal(ctx.uploadedRosterSync.preview, null);
    console.log('Roster scan guards and upload actions passed.');
})().catch(error => {console.error(error); process.exitCode = 1;});
