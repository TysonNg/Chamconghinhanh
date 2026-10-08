const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync('static/js/app.js', 'utf8');
async function check(kind) {
    const elements = {};
    const calls = [];
    const context = {
        allProjectsList: [{name: 'EHome3', project_id: 'ehome-id'}],
        excelFilename: 'monthly.xlsx', pdfFilename: 'monthly.pdf',
        document: {getElementById: id => elements[id] ||= {value: id.endsWith('project-select') ? 'EHome3' : '', style: {}}},
        apiPost: async (url, body) => { calls.push({url, body}); return {success: true, task_id: 'test'}; },
        showToast: () => {}, checkExcelProgress: () => {}, checkPDFProgress: () => {},
        openExtractionProject: () => {throw new Error('Asked for project a second time');},
    };
    vm.createContext(context);
    vm.runInContext(source.slice(source.indexOf('function selectedExtractionProject('), source.indexOf('async function extractPDF(')), context);
    const name = kind === 'excel' ? 'extractExcel' : 'extractPDF';
    const end = kind === 'excel' ? 'async function checkExcelProgress(' : 'async function checkPDFProgress(';
    vm.runInContext(source.slice(source.indexOf(`async function ${name}(`), source.indexOf(end)), context);
    await context[name]();
    assert.equal(calls.length, 1);
    assert.equal(calls[0].body.project_id, 'ehome-id');
    assert.equal(calls[0].url, `/api/${kind}/extract`);
}
(async () => {
    await check('excel');
    await check('pdf');
    console.log('Selected extraction projects passed for Excel and PDF.');
})().catch(error => {console.error(error); process.exitCode = 1;});
