/* Browser-local, bounded scan journal. Works with the existing SSE + scan APIs. */
class ScanLogStore {
    constructor(entries = []) {
        this.entries = entries.filter(e => e && typeof e.message === 'string' && typeof e.time === 'string')
            .slice(-1000);
        this.states = new Map();
        this.workers = new Map();
        this.slow = new Set();
    }
    add(message, level = 'info', task = null, time = new Date().toISOString()) {
        this.entries.push({time, message: String(message), level,
            task_id: task?.task_id || '', project: task?.project_name || ''});
        this.entries = this.entries.slice(-1000);
    }
    observe(snapshot) {
        const labels = {queued: 'Đang chờ lượt', running: 'Đang quét', cancelling: 'Đang dừng',
            cancelled: 'Đã hủy', completed: 'Hoàn thành', failed: 'Lỗi', interrupted: 'Bị gián đoạn'};
        const activeIds = new Set((snapshot.active || []).map(t => t.task_id));
        for (const task of [...(snapshot.active || []), ...(snapshot.history || [])]) {
            const previous = this.states.get(task.task_id);
            if ((previous || activeIds.has(task.task_id)) && previous?.status !== task.status) {
                const level = task.status === 'failed' ? 'error' : task.status === 'completed' ? 'success' :
                    ['cancelled', 'interrupted'].includes(task.status) ? 'warning' : 'info';
                this.add(`${labels[task.status] || task.status} · ${task.target_folder || ''}${task.errors?.length ? ' · ' + task.errors.join('; ') : ''}`, level, task);
            }
            if (activeIds.has(task.task_id) && (!previous || previous.progress !== task.progress) && task.progress > 0) {
                this.add(`Đã xử lý ${task.progress}/${task.total} nhân viên · Xong gần nhất: ${task.current || '—'}`, 'success', task);
            }
            this.states.set(task.task_id, {status: task.status, progress: task.progress});
            for (const worker of task.active_workers || []) {
                const key = `${task.task_id}|${worker.employee_code || worker.person}`;
                const signature = [worker.stage, worker.day, worker.image, worker.image_index].join('|');
                if (this.workers.get(key) !== signature) {
                    const detail = [worker.person, worker.stage, worker.day ? `Ngày ${worker.day}` : '',
                        worker.image ? `Ảnh ${worker.image_index || '—'}/${worker.image_total || '—'}: ${worker.image}` : ''].filter(Boolean).join(' · ');
                    this.add(detail, 'info', task, worker.updated_at || new Date().toISOString());
                    this.workers.set(key, signature);
                }
                const slowKey = `${key}|${signature}|${worker.updated_at}`;
                if (worker.stage === 'Đang nhận diện ảnh' && worker.idle_seconds >= 90 && !this.slow.has(slowKey)) {
                    this.add(`${worker.person} · ${worker.image || 'Ảnh hiện tại'} đã nhận diện ${worker.idle_seconds} giây; giới hạn mỗi ảnh là 120 giây.`, 'warning', task);
                    this.slow.add(slowKey);
                }
            }
        }
        // Keep bookkeeping bounded, including after many separate scans.
        for (const map of [this.states, this.workers]) while (map.size > 500) map.delete(map.keys().next().value);
        while (this.slow.size > 1000) this.slow.delete(this.slow.values().next().value);
    }
    resultOf(entry) {
        const status = /^\[(matched|no_match|no_face|no_portrait|error|identity_unresolved|not_run)\]/.exec(entry.message)?.[1];
        return status || (entry.level === 'error' ? 'error' : 'other');
    }
    filtered({task = '', level = '', result = '', query = ''} = {}) {
        const normalize = value => value.toLocaleLowerCase('vi-VN').normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/\s+/g, '');
        const needle = normalize(query);
        return this.entries.filter(e => (!task || (task === '__app__' ? !e.task_id : e.task_id === task)) &&
            (!level || e.level === level) && (!result || this.resultOf(e) === result) && (!needle || normalize(`${e.message} ${e.project} ${e.task_id}`).includes(needle)))
            .sort((a, b) => (Date.parse(a.time) || 0) - (Date.parse(b.time) || 0));
    }
}
globalThis.ScanLogStore = ScanLogStore;

(() => {
    const storageKey = 'attendance.scan-journal.v1';
    let saved = [];
    try { saved = JSON.parse(localStorage.getItem(storageKey) || '[]'); } catch (_) { /* Storage is optional. */ }
    const store = new ScanLogStore(Array.isArray(saved) ? saved : []);
    let paused = false;
    let storageAvailable = true;
    const el = id => document.getElementById(id);
    const filters = () => ({task: el('scan-log-task')?.value || '', level: el('scan-log-level')?.value || '', result: el('scan-log-result')?.value || '', query: el('scan-log-search')?.value || ''});
    const timeText = time => {
        const date = new Date(time);
        return Number.isNaN(date.getTime()) ? '' : date.toLocaleTimeString('vi-VN', {hour12: false});
    };
    function render() {
        const output = el('scan-log-output');
        if (!output) return;
        const rows = store.filtered(filters());
        el('scan-log-count').textContent = `${rows.length}/${store.entries.length} dòng${paused ? ' · Đã tạm dừng hiển thị, vẫn thu nhật ký' : ''}${storageAvailable ? '' : ' · Không lưu được trên trình duyệt'}`;
        if (paused) return;
        const levels = {info: 'Thông tin', success: 'Đã xử lý', warning: 'Cảnh báo', error: 'Lỗi'};
        output.innerHTML = rows.slice(-400).map(e => `<div class="scan-log-line scan-log-${escapeHtml(e.level)}">
            <time>${escapeHtml(timeText(e.time))}</time><span class="scan-log-level">${escapeHtml(levels[e.level] || 'Thông tin')}</span>
            <span>${e.task_id ? `<span class="scan-log-scope">${escapeHtml(e.project || e.task_id)} · ${escapeHtml(e.task_id)}</span>` : '<span class="scan-log-scope">Máy chủ</span>'}${escapeHtml(e.message)}</span></div>`).join('') ||
            '<p class="text-hint">Chưa có nhật ký phù hợp. Nhật ký mới sẽ xuất hiện khi hệ thống phát thông báo hoặc thay đổi bước quét.</p>';
        if (el('scan-log-autoscroll').checked) output.scrollTop = output.scrollHeight;
    }
    function update() {
        try { localStorage.setItem(storageKey, JSON.stringify(store.entries)); } catch (_) { storageAvailable = false; }
        const select = el('scan-log-task');
        if (select) {
            const selected = select.value;
            const tasks = new Map(store.entries.filter(e => e.task_id).map(e => [e.task_id, e.project]));
            const signature = JSON.stringify([...tasks]);
            if (select.dataset.tasks !== signature) {
                select.innerHTML = '<option value="">Tất cả lượt / máy chủ</option><option value="__app__">Chỉ thông báo máy chủ</option>' +
                    [...tasks].map(([id, project]) => `<option value="${escapeHtml(id)}">${escapeHtml(project ? project + ' · ' + id : id)}</option>`).join('');
                if ([...select.options].some(option => option.value === selected)) select.value = selected;
                select.dataset.tasks = signature;
            }
        }
        render();
    }
    globalThis.scanLogMonitor = {observe(snapshot) { store.observe(snapshot); update(); }};
    let initialized = false;
    function startScanJournal() {
        if (initialized || !el('scan-log-output')) return;
        initialized = true;
        ['scan-log-task', 'scan-log-level', 'scan-log-result', 'scan-log-search', 'scan-log-autoscroll'].forEach(id => el(id).addEventListener('input', render));
        el('scan-log-pause').addEventListener('click', () => {
            paused = !paused;
            el('scan-log-pause').textContent = paused ? 'Tiếp tục' : 'Tạm dừng';
            render();
        });
        el('scan-log-clear').addEventListener('click', () => { store.entries = []; update(); });
        el('scan-log-download').addEventListener('click', () => {
            const rows = store.filtered(filters()).map(e => `${e.time} [${e.level}] [${e.task_id || 'Máy chủ'}] ${e.project} ${e.message}`);
            const url = URL.createObjectURL(new Blob(['\uFEFF' + rows.join('\r\n')], {type: 'text/plain;charset=utf-8'}));
            const link = document.createElement('a');
            link.href = url; link.download = `nhat-ky-quet-${new Date().toISOString().slice(0, 10)}.txt`;
            link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
        });
        update();
        if (typeof EventSource === 'undefined') {
            el('scan-log-connection').textContent = 'Trình duyệt không hỗ trợ log máy chủ; vẫn theo dõi tiến độ mỗi 2 giây.';
            return;
        }
        const stream = new EventSource('/api/log-stream');
        stream.onopen = () => { el('scan-log-connection').textContent = 'Đã kết nối nhật ký máy chủ'; };
        stream.onerror = () => { el('scan-log-connection').textContent = 'Đang kết nối lại nhật ký máy chủ…'; };
        stream.onmessage = event => {
            try {
                const data = JSON.parse(event.data);
                if (!data.message || data.type === 'heartbeat') return;
                store.add(data.message, ['success', 'warning', 'error'].includes(data.type) ? data.type : 'info');
                update();
            } catch (_) { /* Ignore malformed stream frames. */ }
        };
        globalThis.addEventListener('pagehide', () => stream.close(), {once: true});
    }
    globalThis.startScanJournal = startScanJournal;
    document.addEventListener('DOMContentLoaded', startScanJournal);
})();
