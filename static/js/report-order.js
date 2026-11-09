/* One saved employee sequence per project, shared by Excel and PDF scans. */
let reportOrderEmployees = [];
let reportOrderProject = '';
let reportOrderDirty = false;
let reportOrderSaving = false;
let reportOrderRequest = 0;
const employeePauseSaving = new Set();

function renderReportOrder() {
    const list = document.getElementById('report-order-list');
    list.innerHTML = reportOrderEmployees.map((employee, index) => {
        const name = escapeHtml(employee.display_name);
        const code = escapeHtml(employee.memberships?.[0]?.payroll_code || employee.internal_code || 'Chưa có mã');
        return `<li draggable="${!reportOrderSaving}" data-order-index="${index}" class="report-order-item">
            <span class="report-order-grip" aria-hidden="true">⠿</span>
            <input class="form-input report-order-position" type="number" min="1" max="${reportOrderEmployees.length}" value="${index + 1}"
                aria-label="Vị trí của ${name}" data-order-position="${index}" ${reportOrderSaving ? 'disabled' : ''}>
            <span class="report-order-person"><strong>${name}</strong><span class="text-hint">Mã: ${code}${employee.active ? '' : ' · Đã lưu trữ'}</span>${employee.scan_paused ? '<span class="scan-paused-badge">⏸ Tạm nghỉ · Bỏ qua quét và báo cáo</span>' : ''}</span>
            <span class="report-order-actions">
                <button class="btn btn-secondary btn-sm" type="button" data-order-pause="${index}" aria-label="${employee.scan_paused ? 'Cho quét lại' : 'Tạm nghỉ'}: ${name}" ${reportOrderSaving || employeePauseSaving.has(`${reportOrderProject}:${employee.employee_id}`) ? 'disabled' : ''}>${employee.scan_paused ? '▶ Cho quét lại' : '⏸ Tạm nghỉ'}</button>
                <button class="btn btn-secondary btn-sm" type="button" data-order-move="${index}" data-direction="-1" aria-label="Đưa ${name} lên" ${index === 0 || reportOrderSaving ? 'disabled' : ''}>↑</button>
                <button class="btn btn-secondary btn-sm" type="button" data-order-move="${index}" data-direction="1" aria-label="Đưa ${name} xuống" ${index === reportOrderEmployees.length - 1 || reportOrderSaving ? 'disabled' : ''}>↓</button>
            </span></li>`;
    }).join('');
    document.getElementById('report-order-save').disabled = !reportOrderDirty || reportOrderSaving || employeePauseSaving.size > 0;
    document.getElementById('report-order-project').disabled = reportOrderSaving;
}

function moveReportEmployee(from, to) {
    if (reportOrderSaving || !Number.isInteger(from) || !Number.isInteger(to) ||
        from < 0 || from >= reportOrderEmployees.length || to < 0 || to >= reportOrderEmployees.length || from === to) return;
    const [employee] = reportOrderEmployees.splice(from, 1);
    reportOrderEmployees.splice(to, 0, employee);
    reportOrderDirty = true;
    document.getElementById('report-order-state').textContent = 'Có thay đổi chưa lưu. Bấm “Lưu thứ tự” để áp dụng cho lượt quét sau.';
    renderReportOrder();
    document.querySelector(`[data-order-position="${to}"]`)?.focus();
}

async function loadReportOrder() {
    const select = document.getElementById('report-order-project');
    if (reportOrderSaving) { select.value = reportOrderProject; return; }
    if (reportOrderDirty && !await confirmAction('Thứ tự đang sửa chưa được lưu. Bỏ thay đổi và nạp lại danh sách?')) {
        select.value = reportOrderProject;
        return;
    }
    const project = select.value;
    const requestId = ++reportOrderRequest;
    reportOrderProject = project;
    reportOrderDirty = false;
    reportOrderEmployees = [];
    renderReportOrder();
    const state = document.getElementById('report-order-state');
    if (!project) { state.textContent = 'Chưa có dự án để sắp xếp.'; return; }
    state.textContent = 'Đang tải danh sách nhân viên…';
    try {
        const result = await apiGet(`/api/projects/report-order?project_id=${encodeURIComponent(project)}`);
        if (requestId !== reportOrderRequest) return;
        if (!result.success) throw new Error(result.error || 'Không tải được thứ tự nhân viên');
        reportOrderEmployees = result.employees;
        state.textContent = reportOrderEmployees.length ? `${reportOrderEmployees.length} nhân viên · Thứ tự đang dùng cho báo cáo.` : 'Dự án chưa có nhân viên. Thêm hoặc đồng bộ nhân viên trước khi sắp xếp.';
        renderReportOrder();
    } catch (error) {
        if (requestId === reportOrderRequest) state.textContent = `Không tải được danh sách: ${error.message}`;
    }
}

async function saveReportOrder() {
    if (reportOrderSaving || employeePauseSaving.size || !reportOrderDirty || !reportOrderProject) return;
    reportOrderSaving = true;
    renderReportOrder();
    try {
        const result = await apiPost('/api/projects/report-order', {project_id: reportOrderProject,
            employee_ids: reportOrderEmployees.map(employee => employee.employee_id)});
        if (!result.success) throw new Error(result.error || 'Không lưu được thứ tự');
        reportOrderEmployees = result.employees;
        reportOrderDirty = false;
        document.getElementById('report-order-state').textContent = 'Đã lưu. Các lượt quét Excel/PDF tiếp theo sẽ dùng thứ tự này.';
        showToast('Đã lưu thứ tự nhân viên cho dự án.', 'success');
    } catch (error) {
        document.getElementById('report-order-state').textContent = `Chưa lưu được: ${error.message}. Thứ tự đang sửa được giữ lại để thử lại.`;
        showToast(error.message, 'error');
    } finally {
        reportOrderSaving = false;
        renderReportOrder();
    }
}

async function setEmployeeScanPaused(employeeId, paused, project, button) {
    const targetProject = typeof allProjectsList !== 'undefined'
        ? allProjectsList.find(p => p.project_id === project || p.name === project)?.project_id || project : project;
    const key = `${targetProject}:${employeeId}`;
    if (employeePauseSaving.has(key) || reportOrderSaving) return;
    employeePauseSaving.add(key);
    if (button) button.disabled = true;
    renderReportOrder();
    try {
        const result = await apiPost('/api/portraits/employee/scan-pause', {project: targetProject, employee_id: employeeId, paused});
        if (!result.success) throw new Error(result.error || 'Không lưu được trạng thái tạm nghỉ');
        if (result.project_id === reportOrderProject) {
            const employee = reportOrderEmployees.find(e => e.employee_id === employeeId);
            if (employee) employee.scan_paused = result.scan_paused;
        }
        if (typeof allEmployeesList !== 'undefined' && typeof allProjectsList !== 'undefined' &&
            (currentProjectName === project || allProjectsList.some(p => p.name === currentProjectName && p.project_id === result.project_id))) {
            const employee = allEmployeesList.find(e => e.employee_id === employeeId);
            if (employee) employee.scan_paused = result.scan_paused;
            renderEmployeeCards(allEmployeesList);
        }
        showToast(paused ? 'Đã đánh dấu tạm nghỉ. Các lượt quét sau sẽ bỏ qua nhân viên này và không đưa vào báo cáo.' : 'Đã cho phép quét và đưa nhân viên vào báo cáo trở lại.', 'success');
    } catch (error) {
        showToast(error.message, 'error');
    } finally {
        employeePauseSaving.delete(key);
        if (button) button.disabled = false;
        renderReportOrder();
    }
}

function openReportOrderSettings(projectName) {
    navigateToResults('summary');
    const project = allProjectsList.find(p => p.name === projectName);
    const select = document.getElementById('report-order-project');
    if (project && project.project_id !== select.value) { select.value = project.project_id; loadReportOrder(); }
    else if (!reportOrderDirty) loadReportOrder();
    document.getElementById('report-order-section')?.scrollIntoView({block: 'start', behavior: 'smooth'});
}

globalThis.reportOrderEditor = {setProjects(projects, currentProject) {
    const select = document.getElementById('report-order-project');
    if (!select || reportOrderSaving) return;
    const previous = select.value;
    select.innerHTML = projects.map(project => `<option value="${escapeHtml(project.project_id)}">${escapeHtml(project.display_name || project.name)}</option>`).join('');
    const selected = projects.find(project => project.project_id === previous) || projects.find(project => project.name === currentProject) || projects[0];
    select.value = selected?.project_id || '';
    if (select.value !== reportOrderProject || !reportOrderProject) loadReportOrder();
}};

document.addEventListener('DOMContentLoaded', () => {
    const list = document.getElementById('report-order-list');
    if (!list) return;
    let dragged = null;
    list.addEventListener('click', event => {
        const pause = event.target.closest('[data-order-pause]');
        if (pause) {
            const employee = reportOrderEmployees[Number(pause.dataset.orderPause)];
            setEmployeeScanPaused(employee.employee_id, !employee.scan_paused, reportOrderProject, pause);
            return;
        }
        const button = event.target.closest('[data-order-move]');
        if (button) moveReportEmployee(Number(button.dataset.orderMove), Number(button.dataset.orderMove) + Number(button.dataset.direction));
    });
    list.addEventListener('change', event => {
        if (event.target.dataset.orderPosition === undefined) return;
        moveReportEmployee(Number(event.target.dataset.orderPosition), Number(event.target.value) - 1);
        renderReportOrder();
    });
    list.addEventListener('dragstart', event => {
        if (reportOrderSaving || event.target.closest('input, button')) { event.preventDefault(); return; }
        const row = event.target.closest('[data-order-index]');
        if (!row) return;
        dragged = Number(row.dataset.orderIndex);
        event.dataTransfer.effectAllowed = 'move';
        event.dataTransfer.setData('text/plain', String(dragged));
        row.classList.add('dragging');
    });
    list.addEventListener('dragover', event => { if (dragged !== null) event.preventDefault(); });
    list.addEventListener('drop', event => {
        const row = event.target.closest('[data-order-index]');
        if (row && dragged !== null) { event.preventDefault(); moveReportEmployee(dragged, Number(row.dataset.orderIndex)); }
        dragged = null;
    });
    list.addEventListener('dragend', () => { dragged = null; list.querySelector('.dragging')?.classList.remove('dragging'); });
});
