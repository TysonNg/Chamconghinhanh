/* Shared Excel/PDF scan monitor, restored from server state on every page load. */
let scanTasksLoading = false;
const scanTaskActions = new Set();
let scanAIConfigSaving = false;
const scanTaskPreviousStates = new Map();

function openScanDrawer() {
    const drawer = document.getElementById('scan-drawer');
    if (!drawer || drawer.open) return;
    drawer.showModal();
    document.body.classList.add('scan-drawer-open');
    document.getElementById('scan-status-button').setAttribute('aria-expanded', 'true');
    refreshScanTasks();
}

function closeScanDrawer() {
    document.getElementById('scan-drawer')?.close();
}

function initializeScanDrawer() {
    const drawer = document.getElementById('scan-drawer');
    drawer?.addEventListener('click', event => {
        if (event.target === drawer) closeScanDrawer();
    });
    drawer?.addEventListener('close', () => {
        document.body.classList.remove('scan-drawer-open');
        document.getElementById('scan-status-button').setAttribute('aria-expanded', 'false');
    });
}

function renderScanStatus(result) {
    const button = document.getElementById('scan-status-button');
    const label = document.getElementById('scan-status-text');
    if (!button || !label) return;
    const c = result.counts;
    const running = result.active.filter(task => task.status === 'running');
    const total = running.reduce((sum, task) => sum + (Number(task.total) || 0), 0);
    const done = running.reduce((sum, task) => sum + (Number(task.progress) || 0), 0);
    const percent = ` ${total > 0 ? Math.max(0, Math.min(100, Math.round(done / total * 100))) : 0}%`;
    const state = c.cancelling ? 'Đang dừng' : c.running ? `Đang quét${percent}` : c.queued ? 'Đang chờ' : 'Đang nghỉ';
    label.textContent = `Quét mặt · ${state}${c.queued ? ` · ${c.queued} chờ` : ''}`;
    button.dataset.state = c.cancelling || c.running || c.queued ? 'active' : 'idle';
    button.title = 'Mở quản lý quét mặt, tiến độ và nhật ký';
}

function notifyScanTaskChanges(result) {
    const tasks = [...result.active, ...result.history];
    for (const task of tasks) {
        const previous = scanTaskPreviousStates.get(task.task_id);
        if (previous && previous !== task.status && ['completed', 'failed', 'interrupted'].includes(task.status)) {
            const state = task.status === 'completed' ? 'Đã quét xong' : task.status === 'failed' ? 'Quét gặp lỗi' : 'Quét bị gián đoạn';
            showToast(`${state}: ${task.project_name || task.target_folder || task.task_id}`,
                task.status === 'completed' ? 'success' : 'error', 'Xem chi tiết', () => {
                    openScanDrawer();
                    const history = document.getElementById('scan-task-history-details');
                    if (history) history.open = true;
                });
        }
        scanTaskPreviousStates.set(task.task_id, task.status);
    }
    const ids = new Set(tasks.map(task => task.task_id));
    for (const id of scanTaskPreviousStates.keys()) {
        if (!ids.has(id)) scanTaskPreviousStates.delete(id);
    }
}

function renderScanAIConfig(config) {
    const select = document.getElementById('scan-ai-mode');
    if (!select) return;
    if (!config) {
        select.disabled = true;
        document.getElementById('scan-ai-save').disabled = true;
        document.getElementById('scan-ai-state').textContent = 'Máy chủ đang chạy phiên bản cũ. Mở lại phần mềm để bật cấu hình số lượt AI.';
        return;
    }
    select.disabled = false;
    const signature = JSON.stringify([config.hardware_max, config.mode, config.pending_mode]);
    if (select.dataset.config !== signature) {
        select.innerHTML = `<option value="auto">Tự động (tối đa ${Number(config.hardware_max)} AI)</option>` +
            Array.from({length: Number(config.hardware_max)}, (_, i) => `<option value="${i+1}">${i+1} AI cùng lúc</option>`).join('');
        select.value = String(config.pending_mode ?? config.mode);
        select.dataset.config = signature;
    }
    document.getElementById('scan-ai-save').disabled = scanAIConfigSaving;
    document.getElementById('scan-ai-hardware').textContent =
        `${config.cpu_count} nhân CPU · RAM ${config.total_ram_gb ?? 'chưa rõ'} GB (trống ${config.available_ram_gb ?? 'chưa rõ'} GB) · ${config.threads_per_worker} luồng CPU cho mỗi AI`;
    document.getElementById('scan-ai-state').textContent =
        `AI đang xử lý: ${config.active}/${config.max_workers} · Đang chờ: ${config.waiting}` +
        (config.pending_max_workers ? ` · Sẽ đổi sang ${config.pending_max_workers} AI sau khi các ảnh hiện tại xong.` : '') +
        (config.memory_waiting ? ` · ${config.memory_waiting} yêu cầu đang chờ đủ RAM để mở thêm AI. Có thể đóng bớt ứng dụng khác để giải phóng RAM.` : '');
}

async function saveScanAIConfig() {
    if (scanAIConfigSaving) return;
    const value = document.getElementById('scan-ai-mode').value;
    scanAIConfigSaving = true;
    const button = document.getElementById('scan-ai-save');
    button.disabled = true;
    try {
        const result = await apiPost('/api/tasks/ai-config', {mode: value === 'auto' ? 'auto' : Number(value)});
        if (!result.success) throw new Error(result.error || 'Không lưu được số lượt AI');
        renderScanAIConfig(result.inference);
        showToast(result.inference.pending_max_workers ? 'Đã lưu. Sẽ áp dụng sau khi các ảnh hiện tại xong.' : 'Đã áp dụng số lượt AI.', 'success');
    } catch (error) {
        showToast(error.message, 'error');
    } finally {
        scanAIConfigSaving = false;
        button.disabled = false;
        refreshScanTasks();
    }
}

async function loadScanJournal() {
    const body = document.querySelector('.scan-manager .card-body');
    if (!body) return;
    let host = document.getElementById('scan-journal-host');
    if (!host) {
        host = document.createElement('div');
        host.id = 'scan-journal-host';
        body.appendChild(host);
    }
    try {
        const response = await fetch('/static/scan-journal.html?v=20261016_results');
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        host.innerHTML = await response.text();
        if (!globalThis.scanLogMonitor) {
            await new Promise((resolve, reject) => {
                const script = document.createElement('script');
                script.src = '/static/js/scan-logs.js?v=20261016_results';
                script.onload = resolve;
                script.onerror = reject;
                document.body.appendChild(script);
            });
        }
        globalThis.startScanJournal();
        refreshScanTasks();
    } catch (error) {
        host.textContent = 'Không tải được bảng nhật ký. Hãy tải lại trang để thử lại.';
    }
}

function scanQueueMessage(task) {
    return task.ahead_count > 0
        ? `Đang chờ lượt: có ${task.ahead_count} lượt phía trước`
        : 'Đang chuẩn bị bắt đầu quét…';
}

function scanTaskRow(task) {
    const labels = {running: 'Đang quét', cancelling: 'Đang dừng…', cancelled: 'Đã hủy',
        completed: 'Hoàn thành', failed: 'Lỗi', interrupted: 'Bị gián đoạn'};
    const state = task.status === 'queued' ? scanQueueMessage(task) : (labels[task.status] || task.status);
    const total = Number(task.total) || 0;
    const done = Number(task.progress) || 0;
    const percent = total > 0 ? Math.max(0, Math.min(100, Math.round(done / total * 100))) : 0;
    const id = escapeHtml(task.task_id);
    let action = '';
    if (['running', 'queued'].includes(task.status)) {
        action = `<button class="btn btn-secondary btn-sm" data-scan-action="cancel" data-task-id="${id}">${task.status === 'queued' ? 'Hủy chờ' : 'Dừng quét'}</button>`;
    } else if (['failed', 'cancelled', 'interrupted'].includes(task.status)) {
        action = `<button class="btn btn-primary btn-sm" data-scan-action="retry" data-task-id="${id}">Quét lại</button>`;
    } else if (task.status === 'cancelling') {
        action = 'Chờ dừng sau bước hiện tại';
    }
    const errors = (task.errors || []).join('; ');
    const workers = (task.active_workers || []).map(worker =>
        `<div class="scan-worker"><strong>${escapeHtml(worker.person)}</strong> · ${escapeHtml(worker.stage)}
        ${worker.day ? `<div>Ngày ${escapeHtml(worker.day)}</div>` : ''}
        ${worker.image ? `<div>Ảnh ${Number(worker.image_index) || ''}/${Number(worker.image_total) || ''}: ${escapeHtml(worker.image)}</div>` : ''}
        <div class="text-hint">${Number(worker.ai_slot) > 0 ? `Bộ AI #${Number(worker.ai_slot)} · ` : ''}Bước hiện tại: ${Number(worker.idle_seconds) || 0} giây</div></div>`).join('');
    return `<tr><td class="scan-cell-source">${task.task_type === 'excel_face' ? 'Excel' : 'PDF'}<div class="text-hint">${id}</div></td>
        <td class="scan-cell-project">${escapeHtml(task.project_name || 'Chưa gán dự án')}<div class="text-hint">${escapeHtml(task.target_folder)}</div></td>
        <td class="scan-cell-status"><span class="scan-state scan-state-${escapeHtml(task.status)}">${escapeHtml(state)}</span>
        ${errors ? `<div class="text-hint scan-error">${escapeHtml(errors)}</div>` : ''}</td>
        <td class="scan-cell-progress">${total ? `<progress max="100" value="${percent}" aria-label="Tiến độ ${id}"></progress><div>${done}/${total} (${percent}%)</div>` : task.status === 'running' ? '<div>0% · Chưa có nhân viên hoàn tất</div>' : '—'}
        ${workers}
        <div class="text-hint">${task.current ? `Đã xong gần nhất: ${escapeHtml(task.current)}` : ''}</div>
        <div class="text-hint">${escapeHtml(task.created_at ? new Date(task.created_at).toLocaleString('vi-VN') : '')}</div></td>
        <td class="scan-cell-actions">${action}</td></tr>`;
}

async function refreshScanTasks() {
    if (scanTasksLoading || !document.getElementById('scan-task-active')) return;
    scanTasksLoading = true;
    const error = document.getElementById('scan-task-error');
    try {
        const result = await apiGet('/api/tasks/scans');
        if (!result.success) throw new Error(result.error || 'Không tải được trạng thái quét');
        globalThis.scanLogMonitor?.observe(result);
        renderScanAIConfig(result.inference);
        const c = result.counts;
        renderScanStatus(result);
        notifyScanTaskChanges(result);
        document.getElementById('scan-task-summary').textContent =
            `${c.running} đang quét · ${c.cancelling} đang dừng · ${c.queued} đang chờ · Tối đa ${result.max_running} đợt quét chạy cùng lúc`;
        document.getElementById('scan-task-active').innerHTML = result.active.map(scanTaskRow).join('') ||
            '<tr><td colspan="5" class="empty-message">Không có lượt quét đang chạy hoặc chờ.</td></tr>';
        document.getElementById('scan-task-history').innerHTML = result.history.map(scanTaskRow).join('') ||
            '<tr><td colspan="5" class="empty-message">Chưa có lịch sử quét.</td></tr>';
        document.querySelectorAll('[data-scan-action]').forEach(btn => {
            btn.disabled = scanTaskActions.has(btn.dataset.taskId);
        });
        error.hidden = true;
    } catch (err) {
        error.textContent = `Chưa cập nhật được trạng thái: ${err.message}. Dữ liệu đang hiển thị có thể đã cũ.`;
        error.hidden = false;
        const button = document.getElementById('scan-status-button');
        const label = document.getElementById('scan-status-text');
        if (button && label) {
            button.dataset.state = 'error';
            label.textContent = 'Quét mặt · Mất kết nối';
            button.title = error.textContent;
        }
    } finally {
        scanTasksLoading = false;
    }
}

async function manageScanTask(action, taskId) {
    if (!['cancel', 'retry'].includes(action) || scanTaskActions.has(taskId)) return;
    scanTaskActions.add(taskId);
    document.querySelectorAll('[data-scan-action]').forEach(btn => {
        if (btn.dataset.taskId === taskId) btn.disabled = true;
    });
    try {
        const result = await apiPost(`/api/task/${action}/${encodeURIComponent(taskId)}`);
        if (!result.success) throw new Error(result.error || 'Không thực hiện được thao tác');
        showToast(action === 'retry' ? 'Đã tạo lượt quét lại. Xem ở danh sách đang xử lý.' :
            (result.status === 'cancelled' ? 'Đã hủy lượt chờ.' : 'Đã yêu cầu dừng. Đang chờ bước hiện tại kết thúc.'), 'info');
    } catch (err) {
        showToast(err.message, 'error');
    } finally {
        scanTaskActions.delete(taskId);
        await refreshScanTasks();
        document.querySelectorAll('[data-scan-action]').forEach(btn => {
            if (btn.dataset.taskId === taskId) btn.disabled = false;
        });
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initializeScanDrawer();
    loadScanJournal();
    document.querySelector('.scan-manager')?.addEventListener('click', event => {
        const btn = event.target.closest('[data-scan-action]');
        if (btn) manageScanTask(btn.dataset.scanAction, btn.dataset.taskId);
    });
    refreshScanTasks();
    // Poll serially; a slow request never creates overlapping polls.
    async function poll() {
        if (!document.hidden) await refreshScanTasks();
        setTimeout(poll, 2000);
    }
    setTimeout(poll, 2000);
    document.addEventListener('visibilitychange', () => {
        if (!document.hidden) refreshScanTasks();
    });
});
