/**
 * JavaScript cho phần mềm chấm công v2.0
 * Đơn giản hóa với 3 tab: Phân Tích + Tách PDF + Kết Quả
 */

// ==================== State ====================
let pdfTaskId = null;
let pdfFaceTaskId = null;
let pdfFilename = null;
let logEventSource = null;
let excelTaskId = null;
let excelFilename = null;
let excelFaceTaskId = null;

// Photo Management State
let currentPhotoSubtab = 'daily';
let activeSelectedDay = '01';
let currentProjectName = 'Chung cư Tân Thuận Đông';
let allProjectsList = [];
let allEmployeesList = [];
let activeEmpModalName = null;

const TRASH_ICON_SVG = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="vertical-align:-1px;"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>`;

async function withLoading(btn, asyncFn) {
    if (btn && btn instanceof HTMLElement) {
        btn.classList.add('is-loading');
        btn.disabled = true;
    }
    try {
        return await asyncFn();
    } finally {
        if (btn && btn instanceof HTMLElement) {
            btn.classList.remove('is-loading');
            btn.disabled = false;
        }
    }
}

// ==================== API Functions ====================

async function parseJsonResponse(response) {
    const text = await response.text();
    try {
        const data = JSON.parse(text);
        return response.ok ? data : {...data, success: false, error: data.error || `HTTP ${response.status}`};
    } catch (_) {
        if (!response.ok) {
            return {
                success: false,
                error: response.status === 404
                    ? `Không tìm thấy endpoint (${response.status}). Vui lòng khởi động lại server (tắt terminal python main.py và chạy lại) để nhận API mới.`
                    : `Máy chủ phản hồi mã lỗi ${response.status} (${response.statusText || 'Lỗi server'})`
            };
        }
        return { success: false, error: text.slice(0, 150) };
    }
}

async function apiGet(url) {
    const response = await fetch(url);
    return parseJsonResponse(response);
}

async function apiPost(url, data = {}) {
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
    });
    return parseJsonResponse(response);
}

// ==================== Vietnamese Date & Month Pickers ====================

function setDatePickerValue(elementOrId, val) {
    const el = typeof elementOrId === 'string' ? document.getElementById(elementOrId) : elementOrId;
    if (!el) return;
    if (el._flatpickr) {
        el._flatpickr.setDate(val || '', true);
    } else {
        el.value = val || '';
    }
}

function initVietnameseDatePickers() {
    if (typeof flatpickr === 'undefined') return;

    if (flatpickr.l10ns && flatpickr.l10ns.vn) {
        flatpickr.localize(flatpickr.l10ns.vn);
    }

    // 1. Month picker cho #photo-period (Tháng/Năm tiếng Việt)
    const monthInput = document.getElementById('photo-period');
    if (monthInput && !monthInput._flatpickr && typeof monthSelectPlugin !== 'undefined') {
        const now = new Date();
        const currentVal = monthInput.value || (now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0'));
        flatpickr(monthInput, {
            locale: 'vn',
            plugins: [
                new monthSelectPlugin({
                    shorthand: false,
                    dateFormat: 'Y-m',
                    altFormat: '\\T\\h\\á\\n\\g m/Y',
                    theme: 'light'
                })
            ],
            altInput: true,
            altInputClass: 'form-input vn-month-picker-alt',
            defaultDate: currentVal,
            onChange: function(selectedDates, dateStr, instance) {
                instance.element.value = dateStr;
                instance.element.dispatchEvent(new Event('change', { bubbles: true }));
            }
        });
    }

    // 2. Date pickers cho các ô ngày thông thường (dd/mm/yyyy tiếng Việt)
    const dateInputs = document.querySelectorAll('.vn-date-picker, #zalo-date-from, #zalo-date-to, #supp-date, #tool-same-date-input');
    dateInputs.forEach(el => {
        if (!el || el._flatpickr) return;
        const currentVal = el.value;
        flatpickr(el, {
            locale: 'vn',
            dateFormat: 'Y-m-d',
            altInput: true,
            altInputClass: (el.className || 'form-input') + ' vn-date-picker-alt',
            altFormat: 'd/m/Y',
            defaultDate: currentVal || undefined,
            allowInput: true,
            onChange: function(selectedDates, dateStr, instance) {
                instance.element.value = dateStr;
                instance.element.dispatchEvent(new Event('change', { bubbles: true }));
            }
        });
    });
}

function getAggregateReportSettings(projectFallback) {
    const projectInput = document.getElementById('project-name');
    const reportProjectName = projectInput && projectInput.value.trim()
        ? projectInput.value.trim()
        : projectFallback;

    if (!reportProjectName) {
        showToast('Vui lòng nhập tên dự án cho báo cáo tổng hợp', 'warning');
        return null;
    }
    return {
        report_project_name: reportProjectName
    };
}

// ==================== Toast Notifications ====================

function showToast(message, type = 'info', actionText = null, actionCallback = null) {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    const msgSpan = document.createElement('span');
    msgSpan.textContent = message;
    toast.appendChild(msgSpan);

    if (actionText && typeof actionCallback === 'function') {
        const actionBtn = document.createElement('button');
        actionBtn.className = 'toast-btn';
        actionBtn.textContent = actionText;
        actionBtn.onclick = (e) => {
            e.stopPropagation();
            toast.remove();
            actionCallback();
        };
        toast.appendChild(actionBtn);
    }

    container.appendChild(toast);
    setTimeout(() => {
        if (toast.parentNode) toast.remove();
    }, actionText ? 7000 : 4000);
}

// ==================== Tab Navigation & Sub-tabs ====================

const VALID_TABS = {
    process: ['excel', 'pdf'],
    results: ['excel', 'pdf', 'summary'],
    photos: ['daily', 'portraits'],
    zalo: [],
    supplement: []
};

let currentActiveTab = 'process';
let currentProcessSubtab = 'excel';
let currentResultSubtab = 'excel';

function updateUrlParams(tab, subtab = null, push = false) {
    try {
        const url = new URL(window.location.href);
        const prevTab = url.searchParams.get('tab');
        const prevSubtab = url.searchParams.get('subtab');

        const nextSubtab = subtab || null;

        if (prevTab === tab && prevSubtab === nextSubtab) {
            return;
        }

        if (tab) {
            url.searchParams.set('tab', tab);
        } else {
            url.searchParams.delete('tab');
        }

        if (nextSubtab) {
            url.searchParams.set('subtab', nextSubtab);
        } else {
            url.searchParams.delete('subtab');
        }

        const newPath = url.pathname + url.search + url.hash;
        if (push) {
            window.history.pushState({ tab, subtab: nextSubtab }, '', newPath);
        } else {
            window.history.replaceState({ tab, subtab: nextSubtab }, '', newPath);
        }
    } catch (e) {
        console.warn('Không thể cập nhật params URL:', e);
    }
}

function switchProcessSubtab(subtab, updateUrl = true) {
    currentProcessSubtab = subtab;
    const isExcel = subtab === 'excel';

    const btnExcel = document.getElementById('subtab-btn-excel');
    const btnPdf = document.getElementById('subtab-btn-pdf');
    const contentExcel = document.getElementById('subtab-excel');
    const contentPdf = document.getElementById('subtab-pdf');

    if (btnExcel) btnExcel.classList.toggle('active', isExcel);
    if (btnPdf) btnPdf.classList.toggle('active', !isExcel);
    if (contentExcel) contentExcel.classList.toggle('active', isExcel);
    if (contentPdf) contentPdf.classList.toggle('active', !isExcel);

    if (isExcel) {
        loadExcelUploads();
        loadExcelExtractedFiles();
    } else {
        loadPDFUploads();
        loadPDFExtractedFiles();
    }

    if (updateUrl && currentActiveTab === 'process') {
        updateUrlParams('process', subtab, false);
    }
}

function switchResultSubtab(subtab, updateUrl = true) {
    currentResultSubtab = subtab;

    const btnExcel = document.getElementById('subtab-result-btn-excel');
    const btnPdf = document.getElementById('subtab-result-btn-pdf');
    const btnSummary = document.getElementById('subtab-result-btn-summary');

    const contentExcel = document.getElementById('subtab-result-excel');
    const contentPdf = document.getElementById('subtab-result-pdf');
    const contentSummary = document.getElementById('subtab-result-summary');

    if (btnExcel) btnExcel.classList.toggle('active', subtab === 'excel');
    if (btnPdf) btnPdf.classList.toggle('active', subtab === 'pdf');
    if (btnSummary) btnSummary.classList.toggle('active', subtab === 'summary');

    if (contentExcel) contentExcel.classList.toggle('active', subtab === 'excel');
    if (contentPdf) contentPdf.classList.toggle('active', subtab === 'pdf');
    if (contentSummary) contentSummary.classList.toggle('active', subtab === 'summary');

    if (subtab === 'excel') {
        loadExcelFaceFiles();
    } else if (subtab === 'pdf') {
        loadPDFFaceFiles();
    } else if (subtab === 'summary') {
        loadAggregateReports();
    }

    if (updateUrl && currentActiveTab === 'results') {
        updateUrlParams('results', subtab, false);
    }
}

function switchTab(tabId, targetSubtab = null, pushHistory = true) {
    if (!VALID_TABS.hasOwnProperty(tabId)) {
        tabId = 'process';
    }

    currentActiveTab = tabId;

    document.querySelectorAll('.nav-item').forEach(nav => {
        nav.classList.toggle('active', nav.dataset.tab === tabId);
    });

    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.toggle('active', tab.id === `tab-${tabId}`);
    });

    const titles = {
        process: 'Xử Lý Chấm Công',
        results: 'Báo Cáo & Kết Quả',
        photos: 'Quản Lý Ảnh',
        zalo: 'Đồng Bộ & Tải Ảnh Zalo',
        supplement: 'Bổ Sung Ảnh Chấm Công'
    };
    const titleEl = document.querySelector('.page-title');
    if (titleEl) titleEl.textContent = titles[tabId] || 'Chấm Công';

    let activeSub = null;
    if (tabId === 'process') {
        const sub = (targetSubtab && VALID_TABS.process.includes(targetSubtab)) ? targetSubtab : (currentProcessSubtab || 'excel');
        switchProcessSubtab(sub, false);
        activeSub = sub;
    } else if (tabId === 'results') {
        const sub = (targetSubtab && VALID_TABS.results.includes(targetSubtab)) ? targetSubtab : (currentResultSubtab || 'excel');
        switchResultSubtab(sub, false);
        activeSub = sub;
    } else if (tabId === 'photos') {
        const sub = (targetSubtab && VALID_TABS.photos.includes(targetSubtab)) ? targetSubtab : (currentPhotoSubtab || 'daily');
        if (typeof switchPhotoSubtab === 'function') {
            switchPhotoSubtab(sub, false);
        }
        activeSub = sub;
    } else if (tabId === 'zalo') {
        try {
            if (typeof initZaloTab === 'function') {
                initZaloTab();
            }
        } catch (e) {
            console.error('[TabNav] Lỗi initZaloTab:', e);
        }
    } else if (tabId === 'supplement') {
        try {
            if (typeof supplementLoadRecords === 'function') {
                supplementLoadRecords();
                supplementLoadEmployees();
            }
        } catch (e) {
            console.error('[TabNav] Lỗi supplementLoadRecords:', e);
        }
    }

    updateUrlParams(tabId, activeSub, pushHistory);
}

function navigateToResults(subtab = 'excel') {
    switchTab('results', subtab, true);
}

function navigateToProcessTab() {
    switchTab('process', currentProcessSubtab, true);
}

function initTabFromUrl(pushHistory = false) {
    const params = new URLSearchParams(window.location.search);
    const tabParam = params.get('tab');
    const subtabParam = params.get('subtab');

    console.log('[TabNav] initTabFromUrl:', { tabParam, subtabParam, hasTab: !!tabParam, isValid: tabParam ? !!VALID_TABS[tabParam] : false });

    if (tabParam && VALID_TABS.hasOwnProperty(tabParam)) {
        switchTab(tabParam, subtabParam, pushHistory);
    } else {
        switchTab('process', 'excel', false);
    }
}

// Browser Back / Forward handler
window.addEventListener('popstate', () => {
    initTabFromUrl(false);
});

// Gán sự kiện click cho các tab bên sidebar
document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        const tabId = item.dataset.tab;
        switchTab(tabId, null, true);
    });
});

// Khôi phục tab ngay khi script load (DOM đã sẵn sàng vì script ở cuối body)
initTabFromUrl(false);
// ==================== Results ====================

async function loadAggregateReports(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/aggregate-reports');
            const tbody = document.getElementById('aggregate-reports-table-body');
            if (!tbody) return;

            if (data.files && data.files.length > 0) {
                tbody.innerHTML = data.files.map((file, index) => {
                    return `
                        <tr>
                            <td>${index + 1}</td>
                            <td>${file.name}</td>
                            <td><span class="badge">${file.source}</span></td>
                            <td>${file.folder}</td>
                            <td>${formatFileSize(file.size)}</td>
                            <td>
                                <div style="display:flex;gap:6px;align-items:center;justify-content:center;">
                                    <a href="${file.download_url}" class="btn btn-primary btn-sm">Tải về</a>
                                </div>
                            </td>
                        </tr>
                    `;
                }).join('');
            } else {
                tbody.innerHTML = '<tr><td colspan="6" class="empty-message">Chưa có file giải trình tổng hợp. Hãy chạy quét mặt từ Excel hoặc PDF.</td></tr>';
            }
        } catch (error) {
            console.error('Lỗi tải báo cáo tổng hợp:', error);
        }
    });
}

// ==================== PDF Extraction ====================

// PDF Upload Zone
document.addEventListener('DOMContentLoaded', () => {
    initVietnameseDatePickers();
    const uploadZone = document.getElementById('pdf-upload-zone');
    const fileInput = document.getElementById('pdf-file-input');

    if (uploadZone && fileInput) {
        uploadZone.addEventListener('click', () => fileInput.click());

        uploadZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadZone.classList.add('dragover');
        });

        uploadZone.addEventListener('dragleave', () => {
            uploadZone.classList.remove('dragover');
        });

        uploadZone.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0 && files[0].name.toLowerCase().endsWith('.pdf')) {
                handlePDFFile(files[0]);
            } else {
                showToast('Vui lòng chọn file PDF', 'warning');
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handlePDFFile(e.target.files[0]);
            }
        });
    }

    loadAggregateReports();
});

async function handlePDFFile(file) {
    showToast('Đang upload file PDF...', 'info');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/pdf/upload', {
            method: 'POST',
            body: formData
        });
        const result = await response.json();

        if (result.success) {
            pdfFilename = result.filename;
            document.getElementById('pdf-filename').textContent = result.filename;
            document.getElementById('pdf-selected-file').style.display = 'block';
            showToast(`Đã upload: ${result.filename}`, 'success');
            loadPDFUploads();
            await syncUploadedRoster(result.filename, 'pdf');
        } else {
            showToast(result.error || 'Lỗi upload', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

function scanMonthControls() {
    const year = new Date().getFullYear();
    return `<label style="font-size:12px">Tháng quét <select class="face-scan-month form-input" aria-label="Tháng quét" style="width:130px;display:inline-block"><option value="">Theo bảng</option>${Array.from({length:12},(_,i) => `<option value="${String(i+1).padStart(2,'0')}">Tháng ${i+1}</option>`).join('')}</select></label><label style="font-size:12px">Năm <select class="face-scan-year form-input" aria-label="Năm quét" style="width:95px;display:inline-block">${Array.from({length:31},(_,i) => year-20+i).map(y => `<option value="${y}" ${y===year ? 'selected' : ''}>${y}</option>`).join('')}</select></label>`;
}
let extractionProjectFolder = null;
async function assignBatchProject(folder, kind) {
    extractionProjectKind = kind;
    extractionProjectFolder = folder;
    document.getElementById('extract-project-filename').textContent = folder;
    const select = document.getElementById('extract-project-select');
    select.replaceChildren(new Option('-- Chọn dự án của đợt tách --',''));
    allProjectsList.forEach(p => select.add(new Option(p.name,p.project_id)));
    document.getElementById('extract-project-confirm').textContent = 'Lưu dự án';
    openModal('modal-extract-project');
}
let extractionProjectKind = null;
function openExtractionProject(kind) {
    const filename = kind === 'excel' ? excelFilename : pdfFilename;
    if (!filename) return showToast('Chọn file trước khi tách', 'warning');
    extractionProjectKind = kind;
    extractionProjectFolder = null;
    document.getElementById('extract-project-confirm').textContent = 'Tách file';
    document.getElementById('extract-project-filename').textContent = filename;
    const select = document.getElementById('extract-project-select');
    select.replaceChildren(new Option('-- Chọn dự án của bảng chấm công --',''));
    allProjectsList.forEach(p => select.add(new Option(p.name,p.project_id)));
    openModal('modal-extract-project');
}
async function confirmProjectExtraction() {
    const projectId = document.getElementById('extract-project-select').value;
    if (!projectId) return showToast('Chọn dự án trước khi tách', 'warning');
    const project = allProjectsList.find(p => p.project_id === projectId);
    const select = document.getElementById(`${extractionProjectKind}-project-select`);
    if (select && project) select.value = project.name;
    if (extractionProjectFolder) {
        const res = await apiPost('/api/extraction/project', {kind: extractionProjectKind, folder: extractionProjectFolder, project_id: projectId});
        if (!res.success) return showToast(res.error, 'error');
        closeModal('modal-extract-project');
        await loadExcelExtractedFiles(); await loadPDFExtractedFiles();
        return;
    }
    closeModal('modal-extract-project');
    if (extractionProjectKind === 'excel') await extractExcel(projectId);
    else await extractPDF(projectId);
}

function selectedExtractionProject(kind) {
    const name = document.getElementById(`${kind}-project-select`)?.value;
    return allProjectsList.find(p => p.name === name)?.project_id || null;
}

async function extractPDF(projectId = null) {
    projectId = projectId || selectedExtractionProject('pdf');
    if (!projectId) return openExtractionProject('pdf');
    if (!pdfFilename) {
        showToast('Vui lòng chọn file PDF trước', 'warning');
        return;
    }

    const btn = document.getElementById('btn-extract');
    const progressSection = document.getElementById('pdf-progress-section');

    btn.disabled = true;
    btn.textContent = ' Đang xử lý...';
    progressSection.style.display = 'block';

    try {
        const result = await apiPost('/api/pdf/extract', { filename: pdfFilename, project_id: projectId });

        if (result.success) {
            pdfTaskId = result.task_id;
            showToast('Đã bắt đầu tách PDF...', 'info');
            checkPDFProgress();
        } else {
            showToast(result.error || 'Lỗi tách PDF', 'error');
            btn.disabled = false;
            btn.textContent = ' Bắt Đầu Tách';
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
        btn.disabled = false;
        btn.textContent = ' Bắt Đầu Tách';
    }
}

async function checkPDFProgress() {
    if (!pdfTaskId) return;

    try {
        const result = await apiGet(`/api/pdf/status/${pdfTaskId}`);

        // Update progress
        document.getElementById('pdf-progress-title').textContent = result.message || 'Đang xử lý...';
        document.getElementById('pdf-progress-percent').textContent = `${result.progress}%`;
        document.getElementById('pdf-progress-fill').style.width = `${result.progress}%`;
        document.getElementById('pdf-progress-detail').textContent =
            result.current_page ? `Trang ${result.current_page}/${result.total}` : '';

        if (result.status === 'completed') {
            showToast(`Hoàn thành! Đã tạo ${result.files_created.length} file Word.`, 'success');
            document.getElementById('btn-extract').disabled = false;
            document.getElementById('btn-extract').textContent = ' Bắt Đầu Tách';
            loadPDFExtractedFiles();

            setTimeout(() => {
                document.getElementById('pdf-progress-section').style.display = 'none';
            }, 2000);
        } else if (result.status === 'error') {
            showToast('Lỗi: ' + result.error, 'error');
            document.getElementById('btn-extract').disabled = false;
            document.getElementById('btn-extract').textContent = ' Bắt Đầu Tách';
        } else {
            // Continue checking
            setTimeout(checkPDFProgress, 1000);
        }
    } catch (error) {
        console.error('Lỗi kiểm tra tiến độ:', error);
        setTimeout(checkPDFProgress, 2000);
    }
}

async function loadPDFUploads(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/pdf/uploads');
            const tbody = document.getElementById('pdf-uploads-table-body');

            if (data.files && data.files.length > 0) {
                document.getElementById('stat-pdf-uploads').textContent = data.files.length;
                tbody.innerHTML = data.files.map((file, index) => {
                    const safeName = file.name.replace(/'/g, "\\'");
                    return `
                        <tr>
                            <td>${index + 1}</td>
                            <td>${file.name}</td>
                            <td>${formatFileSize(file.size)}</td>
                            <td>
                                <div style="display:flex;gap:6px;align-items:center;justify-content:center;">
                                    <button class="btn btn-success btn-sm" onclick="selectPDFForExtract('${safeName}')">Tách</button>
                                    <button class="btn btn-secondary btn-sm" onclick="syncUploadedRoster('${safeName}', 'pdf')">Đồng bộ nhân viên</button>
                                    <button class="btn btn-icon-danger btn-sm" title="Xóa file PDF này" onclick="deletePDFUpload('${safeName}')">
                                        ${TRASH_ICON_SVG}<span>Xóa</span>
                                    </button>
                                </div>
                            </td>
                        </tr>
                    `;
                }).join('');
            } else {
                document.getElementById('stat-pdf-uploads').textContent = '0';
                tbody.innerHTML = '<tr><td colspan="4" class="empty-message">Chưa có file PDF nào</td></tr>';
            }
        } catch (error) {
            console.error('Lỗi load PDF uploads:', error);
        }
    });
}

function selectPDFForExtract(filename) {
    pdfFilename = filename;
    document.getElementById('pdf-filename').textContent = filename;
    document.getElementById('pdf-selected-file').style.display = 'block';
    return openExtractionProject('pdf');
}

async function loadPDFExtractedFiles(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/pdf/files');
            const container = document.getElementById('pdf-extracted-files');

            if (data.folders && data.folders.length > 0) {
                document.getElementById('stat-pdf-extracted').textContent = data.folders.length;

                let html = '';
                for (const folder of data.folders) {
                    const safeFolder = folder.folder.replace(/'/g, "\\'");
                    html += `
                        <div class="card" style="margin-bottom: 16px;">
                            <div class="card-header">
                                <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                                    <h3 style="margin:0;">${escapeHtml(folder.source_filename ? folder.source_filename.replace(/\.[^.]+$/, '') : folder.folder)} (${folder.count} files) · ${escapeHtml(folder.project_name || 'Chưa gán dự án')}</h3>
                                    ${scanMonthControls()}
                                    <button class="btn btn-primary btn-sm" onclick="startPDFFaceAnalyze('${safeFolder}', this, '${folder.project_id || ''}')">Quét mặt</button>
                                    <button class="btn btn-secondary btn-sm" onclick="assignBatchProject('${safeFolder}', 'pdf')">${folder.project_id ? 'Đổi dự án' : 'Gán dự án'}</button>
                                    <button class="btn btn-icon-danger btn-sm" title="Xóa đợt này" onclick="deletePDFExtractedFolder('${safeFolder}')">
                                        ${TRASH_ICON_SVG}<span>Xóa đợt</span>
                                    </button>
                                </div>
                            </div>
                            <div class="card-body">
                                <div class="table-container">
                                    <table class="data-table">
                                        <thead>
                                            <tr>
                                                <th>STT</th>
                                                <th>Tên File</th>
                                                <th>Kích Thước</th>
                                                <th>Thao Tác</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${folder.files.map((file, idx) => {
                                                const safeFile = file.name.replace(/'/g, "\\'");
                                                return `
                                                    <tr>
                                                        <td>${idx + 1}</td>
                                                        <td>${file.name}</td>
                                                        <td>${formatFileSize(file.size)}</td>
                                                        <td>
                                                            <div style="display:flex;gap:6px;align-items:center;">
                                                                <a href="/api/pdf/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}" 
                                                                   class="btn btn-primary btn-sm">Tải về</a>
                                                                <button class="btn btn-icon-danger btn-sm" title="Xóa file này" onclick="deletePDFExtractedFile('${safeFolder}', '${safeFile}')">
                                                                    ${TRASH_ICON_SVG}
                                                                </button>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                `;
                                            }).join('')}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    `;
                }
                container.innerHTML = html;
            } else {
                document.getElementById('stat-pdf-extracted').textContent = '0';
                container.innerHTML = '<p class="empty-message">Chưa có file nào được tách</p>';
            }
        } catch (error) {
            console.error('Lỗi load extracted files:', error);
        }
    });
}

async function startPDFFaceAnalyze(folder, button = null, batchProjectId = '') {
    const batchProject = allProjectsList.find(p => p.project_id === batchProjectId);
    if (batchProject) document.getElementById('pdf-project-select').value = batchProject.name;
    if (!folder) {
        showToast('Vui lòng chọn thư mục đã tách', 'warning');
        return;
    }

    const thresholdInput = document.getElementById('pdf-face-threshold');
    const threshold = thresholdInput ? thresholdInput.value.trim() : '';
    const projSelect = document.getElementById('pdf-project-select');
    const project = projSelect && projSelect.value ? projSelect.value : currentProjectName;
    const reportSettings = getAggregateReportSettings(project);
    if (!reportSettings) return;
    const scanControls = button?.closest('div');
    const monthValue = scanControls?.querySelector('.face-scan-month')?.value || '';
    const yearValue = scanControls?.querySelector('.face-scan-year')?.value || '';
    const scanMonth = monthValue && yearValue ? `${yearValue}-${monthValue}` : '';
    if (scanMonth) reportSettings.scan_month = scanMonth;

    if (!await checkRosterBeforeScan(folder, 'pdf', project)) return;

    const progressSection = document.getElementById('pdf-face-progress-section');
    progressSection.style.display = 'block';
    document.getElementById('pdf-face-progress-fill').style.width = '0%';
    document.getElementById('pdf-face-progress-percent').textContent = '0%';
    document.getElementById('pdf-face-progress-title').textContent = 'Đang khởi động quét mặt từ PDF...';
    document.getElementById('pdf-face-progress-detail').textContent = '';

    try {
        const result = await apiPost('/api/pdf/face/analyze', {
            folder,
            distance_threshold: threshold,
            project: project,
            ...reportSettings
        });
        if (result.success) {
            pdfFaceTaskId = result.task_id;
            const btnCancel = document.getElementById('btn-cancel-pdf-face');
            if (btnCancel) {
                btnCancel.style.display = 'inline-block';
                btnCancel.disabled = false;
                btnCancel.textContent = 'Dừng quét';
            }
            if (result.status === 'queued') {
                showToast('Đã thêm vào hàng đợi (đang có đợt khác quét)...', 'info');
                const badge = document.getElementById('pdf-face-queue-badge');
                if (badge) badge.style.display = 'inline-block';
                document.getElementById('pdf-face-progress-title').textContent = 'Đang chờ lượt...';
            } else {
                showToast('Đã bắt đầu quét mặt từ PDF...', 'info');
                const badge = document.getElementById('pdf-face-queue-badge');
                if (badge) badge.style.display = 'none';
            }
            checkPDFFaceProgress();
        } else {
            showToast(result.error || 'Lỗi quét mặt từ PDF', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

async function checkPDFFaceProgress() {
    if (!pdfFaceTaskId) return;
    try {
        const result = await apiGet(`/api/pdf/face/status/${pdfFaceTaskId}`);
        const pct = result.total > 0 ? Math.round((result.progress / result.total) * 100) : 0;
        const queueBadge = document.getElementById('pdf-face-queue-badge');
        const btnCancel = document.getElementById('btn-cancel-pdf-face');

        if (result.status === 'queued') {
            if (queueBadge) queueBadge.style.display = 'inline-block';
            if (btnCancel) {
                btnCancel.style.display = 'inline-block';
                btnCancel.disabled = false;
                btnCancel.textContent = 'Hủy chờ';
            }
            document.getElementById('pdf-face-progress-title').textContent = 'Đang chờ lượt (có đợt khác đang quét)...';
            document.getElementById('pdf-face-progress-percent').textContent = '0%';
            document.getElementById('pdf-face-progress-fill').style.width = '0%';
            document.getElementById('pdf-face-progress-detail').textContent = 'Hàng đợi tuần tự';
            setTimeout(checkPDFFaceProgress, 1500);
            return;
        }

        if (queueBadge) queueBadge.style.display = 'none';

        if (result.status === 'running') {
            if (btnCancel) {
                btnCancel.style.display = 'inline-block';
                btnCancel.disabled = false;
                btnCancel.textContent = 'Dừng quét';
            }
            document.getElementById('pdf-face-progress-title').textContent =
                result.current ? `Đang xuất: ${result.current}` : 'Đang quét mặt từ PDF...';
            document.getElementById('pdf-face-progress-percent').textContent = `${pct}%`;
            document.getElementById('pdf-face-progress-fill').style.width = `${pct}%`;
            document.getElementById('pdf-face-progress-detail').textContent =
                `${result.progress}/${result.total} file`;
            setTimeout(checkPDFFaceProgress, 1000);
            return;
        }

        if (result.status === 'cancelling') {
            if (btnCancel) {
                btnCancel.disabled = true;
                btnCancel.textContent = 'Đang dừng...';
            }
            document.getElementById('pdf-face-progress-title').textContent = 'Đang dừng quét mặt PDF...';
            setTimeout(checkPDFFaceProgress, 1000);
            return;
        }

        if (result.status === 'cancelled') {
            if (btnCancel) btnCancel.style.display = 'none';
            document.getElementById('pdf-face-progress-title').textContent = 'Đã dừng quét!';
            document.getElementById('pdf-face-progress-detail').textContent = `Đã giữ lại ${result.progress || 0} file Word hoàn tất`;
            showToast(`Đã dừng quét! Đã bảo toàn ${result.progress || 0} file Word cá nhân.`, 'warning');
            loadPDFFaceFiles();
            loadAggregateReports();
            setTimeout(() => {
                document.getElementById('pdf-face-progress-section').style.display = 'none';
            }, 3500);
            return;
        }

        if (result.status === 'interrupted') {
            if (btnCancel) btnCancel.style.display = 'none';
            document.getElementById('pdf-face-progress-title').textContent = 'Đợt quét bị gián đoạn do máy chủ tắt ngang';
            showToast('Đợt quét bị gián đoạn do máy chủ tắt ngang. Bạn có thể bấm quét lại.', 'warning');
            loadPDFFaceFiles();
            setTimeout(() => {
                document.getElementById('pdf-face-progress-section').style.display = 'none';
            }, 4000);
            return;
        }

        if (result.status === 'completed') {
            if (btnCancel) btnCancel.style.display = 'none';
            document.getElementById('pdf-face-progress-percent').textContent = '100%';
            document.getElementById('pdf-face-progress-fill').style.width = '100%';
            showToast(`Hoàn thành! Đã tạo file Word theo nhân viên và báo cáo giải trình tổng hợp.`, 'success', 'Xem Kết Quả Ngay', () => navigateToResults('summary'));
            loadPDFFaceFiles();
            loadAggregateReports();
            setTimeout(() => {
                document.getElementById('pdf-face-progress-section').style.display = 'none';
            }, 3000);
            return;
        }

        if (result.status === 'failed') {
            if (btnCancel) btnCancel.style.display = 'none';
            showToast((result.errors && result.errors[0]) || 'Lỗi quét mặt từ PDF', 'error');
        }
    } catch (error) {
        console.error('Lỗi kiểm tra tiến độ PDF face:', error);
        setTimeout(checkPDFFaceProgress, 2000);
    }
}

async function loadPDFFaceFiles(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/pdf/face/files');
            const container = document.getElementById('pdf-face-files');

            if (data.folders && data.folders.length > 0) {
                let html = '';
                for (const folder of data.folders) {
                    const safeFolder = folder.folder.replace(/'/g, "\\'");
                    html += `
                        <div class="card" style="margin-bottom: 16px;">
                            <div class="card-header">
                                <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                                    <h3 style="margin:0;">${folder.folder} (${folder.count} files)</h3>
                                    <button class="btn btn-icon-danger btn-sm" title="Xóa đợt kết quả này" onclick="deletePDFFaceFolder('${safeFolder}')">
                                        ${TRASH_ICON_SVG}<span>Xóa đợt</span>
                                    </button>
                                </div>
                            </div>
                            <div class="card-body">
                                <div class="table-container">
                                    <table class="data-table">
                                        <thead>
                                            <tr>
                                                <th>STT</th>
                                                <th>Tên File</th>
                                                <th>Kích Thước</th>
                                                <th>Thao Tác</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${folder.files.map((file, idx) => {
                                                const safeFile = file.name.replace(/'/g, "\\'");
                                                return `
                                                    <tr>
                                                        <td>${idx + 1}</td>
                                                        <td>${file.name}</td>
                                                        <td>${formatFileSize(file.size)}</td>
                                                        <td>
                                                            <div style="display:flex;gap:6px;align-items:center;">
                                                                <a href="/api/pdf/face/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}"
                                                                   class="btn btn-primary btn-sm">Tải về</a>
                                                                <button class="btn btn-icon-danger btn-sm" title="Xóa file này" onclick="deletePDFFaceFile('${safeFolder}', '${safeFile}')">
                                                                    ${TRASH_ICON_SVG}
                                                                </button>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                `;
                                            }).join('')}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    `;
                }
                container.innerHTML = html;
            } else {
                container.innerHTML = '<p class="empty-message">Chưa có file nào được quét mặt</p>';
            }
        } catch (error) {
            console.error('Lỗi load PDF face files:', error);
        }
    });
}

// ==================== Utilities ====================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function refreshData(btn) {
    return withLoading(btn, async () => {
        const activeNav = document.querySelector('.nav-item.active');
        const tabId = activeNav ? activeNav.dataset.tab : (currentActiveTab || 'process');

        if (tabId === 'process') {
            await switchProcessSubtab(currentProcessSubtab, false);
        } else if (tabId === 'results') {
            await switchResultSubtab(currentResultSubtab, false);
        } else if (tabId === 'photos') {
            await switchPhotoSubtab(currentPhotoSubtab, false);
        } else if (tabId === 'zalo') {
            if (typeof initZaloTab === 'function') initZaloTab();
        }
        showToast('Đã làm mới dữ liệu', 'success');
    });
}

// ==================== Log Functions ====================

function addLog(message, type = 'default') {
    const logContent = document.getElementById('log-content');
    if (!logContent) return;

    const now = new Date();
    const time = now.toLocaleTimeString('vi-VN');

    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.innerHTML = `
        <span class="log-time">[${time}]</span>
        <span class="log-message">${message}</span>
    `;
    logContent.appendChild(entry);

    // Auto scroll to bottom
    logContent.scrollTop = logContent.scrollHeight;
}

function clearLogs() {
    const logContent = document.getElementById('log-content');
    if (logContent) {
        logContent.innerHTML = '';
    }
}

function toggleLogPanel() {
    const logContent = document.getElementById('log-content');
    if (logContent) {
        logContent.style.display = logContent.style.display === 'none' ? 'block' : 'none';
    }
}

function startLogStream() {
    if (logEventSource) {
        logEventSource.close();
    }

    try {
        logEventSource = new EventSource('/api/log-stream');

        logEventSource.onopen = () => {
            console.log('SSE Connection opened');
            addLog('🔌 Kết nối log stream...', 'info');
        };

        logEventSource.onmessage = (event) => {
            try {
                const data = JSON.parse(event.data);
                // Filter out heartbeat messages
                if (data.type === 'heartbeat' || !data.message) {
                    return;
                }
                addLog(data.message, data.type || 'default');
            } catch (e) {
                if (event.data && event.data.trim()) {
                    addLog(event.data, 'default');
                }
            }
        };

        logEventSource.onerror = (error) => {
            console.log('SSE Error or connection closed');
        };
    } catch (e) {
        console.error('Failed to create EventSource:', e);
        addLog(' Không thể kết nối log stream', 'error');
    }
}

function stopLogStream() {
    if (logEventSource) {
        logEventSource.close();
        logEventSource = null;
    }
}


// ==================== Excel Extraction ====================

function handleExcelFileSelect(input) {
    if (input.files.length > 0) {
        handleExcelFile(input.files[0]);
    }
}

async function handleExcelFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xls', 'xlsx'].includes(ext)) {
        showToast('Vui lòng chọn file .xls hoặc .xlsx', 'warning');
        return;
    }

    showToast('Đang upload file Excel...', 'info');
    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/excel/upload', { method: 'POST', body: formData });
        const result = await response.json();
        if (result.success) {
            excelFilename = result.filename;
            document.getElementById('excel-filename').textContent = result.filename;
            document.getElementById('excel-selected-file').style.display = 'block';
            showToast(`Đã upload: ${result.filename}`, 'success');
            loadExcelUploads();
            await syncUploadedRoster(result.filename, 'excel');
        } else {
            showToast(result.error || 'Lỗi upload', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

async function extractExcel(projectId = null) {
    projectId = projectId || selectedExtractionProject('excel');
    if (!projectId) return openExtractionProject('excel');
    if (!excelFilename) {
        showToast('Vui lòng chọn file Excel trước', 'warning');
        return;
    }
    const btn = document.getElementById('btn-excel-extract');
    const progressSection = document.getElementById('excel-progress-section');
    btn.disabled = true;
    btn.textContent = ' Đang xử lý...';
    progressSection.style.display = 'block';
    document.getElementById('excel-progress-fill').style.width = '0%';
    document.getElementById('excel-progress-percent').textContent = '0%';
    document.getElementById('excel-progress-title').textContent = 'Đang khởi động...';

    try {
        const result = await apiPost('/api/excel/extract', { filename: excelFilename, project_id: projectId });
        if (result.success) {
            excelTaskId = result.task_id;
            showToast('Đã bắt đầu tách Excel...', 'info');
            checkExcelProgress();
        } else {
            showToast(result.error || 'Lỗi tách Excel', 'error');
            btn.disabled = false;
            btn.textContent = ' Bắt Đầu Tách';
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
        btn.disabled = false;
        btn.textContent = ' Bắt Đầu Tách';
    }
}

async function checkExcelProgress() {
    if (!excelTaskId) return;
    try {
        const result = await apiGet(`/api/excel/status/${excelTaskId}`);
        const pct = result.total > 0 ? Math.round((result.progress / result.total) * 100) : 0;
        document.getElementById('excel-progress-title').textContent =
            result.current ? `Đang xuất: ${result.current}` : 'Đang xử lý...';
        document.getElementById('excel-progress-percent').textContent = `${pct}%`;
        document.getElementById('excel-progress-fill').style.width = `${pct}%`;
        document.getElementById('excel-progress-detail').textContent =
            `${result.progress}/${result.total} người`;

        if (result.status === 'completed') {
            document.getElementById('stat-excel-persons').textContent = result.files.length;
            document.getElementById('stat-excel-absent').textContent = 0;
            showToast(`Hoàn thành! Đã tạo ${result.files.length} file Word in theo từng người.`, 'success');
            document.getElementById('btn-excel-extract').disabled = false;
            document.getElementById('btn-excel-extract').textContent = ' Bắt Đầu Tách';
            loadExcelExtractedFiles();
            setTimeout(() => {
                document.getElementById('excel-progress-section').style.display = 'none';
            }, 3000);
        } else if (result.status === 'failed') {
            showToast('Lỗi: ' + (result.errors[0] || 'Không rõ'), 'error');
            document.getElementById('btn-excel-extract').disabled = false;
            document.getElementById('btn-excel-extract').textContent = ' Bắt Đầu Tách';
        } else {
            setTimeout(checkExcelProgress, 800);
        }
    } catch (error) {
        console.error('Lỗi kiểm tra tiến độ Excel:', error);
        setTimeout(checkExcelProgress, 2000);
    }
}

async function loadExcelUploads(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/excel/uploads');
            const tbody = document.getElementById('excel-uploads-table-body');
            const count = data.files ? data.files.length : 0;
            document.getElementById('stat-excel-uploads').textContent = count;

            if (data.files && data.files.length > 0) {
                tbody.innerHTML = data.files.map((file, index) => {
                    const safeName = file.name.replace(/'/g, "\\'");
                    return `
                        <tr>
                            <td>${index + 1}</td>
                            <td>${file.name}</td>
                            <td>${formatFileSize(file.size)}</td>
                            <td>
                                <div style="display:flex;gap:6px;align-items:center;justify-content:center;">
                                    <button class="btn btn-success btn-sm" onclick="selectExcelForExtract('${safeName}')">Tách</button>
                                    <button class="btn btn-secondary btn-sm" onclick="syncUploadedRoster('${safeName}', 'excel')">Đồng bộ nhân viên</button>
                                    <button class="btn btn-icon-danger btn-sm" title="Xóa file này" onclick="deleteExcelUpload('${safeName}')">
                                        ${TRASH_ICON_SVG}<span>Xóa</span>
                                    </button>
                                </div>
                            </td>
                        </tr>
                    `;
                }).join('');
            } else {
                tbody.innerHTML = '<tr><td colspan="4" class="empty-message">Chưa có file Excel nào</td></tr>';
            }
        } catch (error) {
            console.error('Lỗi load Excel uploads:', error);
        }
    });
}

function selectExcelForExtract(filename) {
    excelFilename = filename;
    document.getElementById('excel-filename').textContent = filename;
    document.getElementById('excel-selected-file').style.display = 'block';
    return openExtractionProject('excel');
}

async function loadExcelExtractedFiles(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/excel/files');
            const container = document.getElementById('excel-extracted-files');
            if (data.folders && data.folders.length > 0) {
                let totalFiles = 0;
                let html = '';
                for (const folder of data.folders) {
                    const safeFolder = folder.folder.replace(/'/g, "\\'");
                    totalFiles += folder.count;
                    html += `
                        <div class="card" style="margin-bottom: 16px;">
                            <div class="card-header">
                                <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                                    <h3 style="margin:0;">${escapeHtml(folder.source_filename ? folder.source_filename.replace(/\.[^.]+$/, '') : folder.folder)} (${folder.count} người) · ${escapeHtml(folder.project_name || 'Chưa gán dự án')}</h3>
                                    ${scanMonthControls()}
                                    <button class="btn btn-primary btn-sm" onclick="startExcelFaceAnalyze('${safeFolder}', this, '${folder.project_id || ''}')">Quét mặt</button>
                                    <button class="btn btn-secondary btn-sm" onclick="assignBatchProject('${safeFolder}', 'excel')">${folder.project_id ? 'Đổi dự án' : 'Gán dự án'}</button>
                                    <button class="btn btn-secondary btn-sm" onclick="openExcelExtractedFolder('${safeFolder}')" title="Mở thư mục chứa các file Word">
                                        <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                                            <path d="M3 5h6l2 2h10v12H3z"></path>
                                        </svg>
                                        <span>Mở thư mục Word</span>
                                    </button>
                                    <button class="btn btn-icon-danger btn-sm" title="Xóa đợt này" onclick="deleteExcelExtractedFolder('${safeFolder}')">
                                        ${TRASH_ICON_SVG}<span>Xóa đợt</span>
                                    </button>
                                </div>
                            </div>
                            <div class="card-body">
                                <div class="table-container">
                                    <table class="data-table">
                                        <thead>
                                            <tr>
                                                <th>STT</th>
                                                <th>Tên File Word</th>
                                                <th>Kích Thước</th>
                                                <th>Thao Tác</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${folder.files.map((file, idx) => {
                                                const safeFile = file.name.replace(/'/g, "\\'");
                                                return `
                                                    <tr>
                                                        <td>${idx + 1}</td>
                                                        <td>${file.name}</td>
                                                        <td>${formatFileSize(file.size)}</td>
                                                        <td>
                                                            <div style="display:flex;gap:6px;align-items:center;">
                                                                <a href="/api/excel/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}"
                                                                   class="btn btn-primary btn-sm">Tải về</a>
                                                                <button class="btn btn-icon-danger btn-sm" title="Xóa file này" onclick="deleteExcelExtractedFile('${safeFolder}', '${safeFile}')">
                                                                    ${TRASH_ICON_SVG}
                                                                </button>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                `;
                                            }).join('')}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    `;
                }
                document.getElementById('stat-excel-persons').textContent = totalFiles;
                container.innerHTML = html;
            } else {
                document.getElementById('stat-excel-persons').textContent = '0';
                container.innerHTML = '<p class="empty-message">Chưa có file Word nào được tạo</p>';
            }
        } catch (error) {
            console.error('Lỗi load Excel extracted files:', error);
        }
    });
}

async function openExcelExtractedFolder(folder) {
    try {
        const response = await fetch('/api/open/folder', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type: 'excel_output', subpath: folder })
        });
        const data = await response.json();
        if (data.success) {
            showToast('Đã mở thư mục chứa file Word', 'success');
        } else {
            showToast('Không thể mở thư mục: ' + data.error, 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

async function startExcelFaceAnalyze(folder, button = null, batchProjectId = '') {
    const batchProject = allProjectsList.find(p => p.project_id === batchProjectId);
    if (batchProject) document.getElementById('excel-project-select').value = batchProject.name;
    if (!folder) {
        showToast('Vui lòng chọn thư mục đã tách', 'warning');
        return;
    }

    const thresholdInput = document.getElementById('excel-face-threshold');
    const threshold = thresholdInput ? thresholdInput.value.trim() : '';
    const projSelect = document.getElementById('excel-project-select');
    const project = projSelect && projSelect.value ? projSelect.value : currentProjectName;
    const reportSettings = getAggregateReportSettings(project);
    if (!reportSettings) return;
    const scanControls = button?.closest('div');
    const monthValue = scanControls?.querySelector('.face-scan-month')?.value || '';
    const yearValue = scanControls?.querySelector('.face-scan-year')?.value || '';
    const scanMonth = monthValue && yearValue ? `${yearValue}-${monthValue}` : '';
    if (scanMonth) reportSettings.scan_month = scanMonth;

    if (!await checkRosterBeforeScan(folder, 'excel', project)) return;

    const progressSection = document.getElementById('excel-face-progress-section');
    progressSection.style.display = 'block';
    document.getElementById('excel-face-progress-fill').style.width = '0%';
    document.getElementById('excel-face-progress-percent').textContent = '0%';
    document.getElementById('excel-face-progress-title').textContent = 'Đang khởi động quét mặt...';
    document.getElementById('excel-face-progress-detail').textContent = '';

    try {
        const result = await apiPost('/api/excel/face/analyze', {
            folder,
            distance_threshold: threshold,
            project: project,
            ...reportSettings
        });
        if (result.success) {
            excelFaceTaskId = result.task_id;
            const btnCancel = document.getElementById('btn-cancel-excel-face');
            if (btnCancel) {
                btnCancel.style.display = 'inline-block';
                btnCancel.disabled = false;
                btnCancel.textContent = 'Dừng quét';
            }
            if (result.status === 'queued') {
                showToast('Đã thêm vào hàng đợi (đang có đợt khác quét)...', 'info');
                const badge = document.getElementById('excel-face-queue-badge');
                if (badge) badge.style.display = 'inline-block';
                document.getElementById('excel-face-progress-title').textContent = 'Đang chờ lượt...';
            } else {
                showToast('Đã bắt đầu quét mặt...', 'info');
                const badge = document.getElementById('excel-face-queue-badge');
                if (badge) badge.style.display = 'none';
            }
            checkExcelFaceProgress();
        } else {
            showToast(result.error || 'Lỗi quét mặt', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

async function checkExcelFaceProgress() {
    if (!excelFaceTaskId) return;
    try {
        const result = await apiGet(`/api/excel/face/status/${excelFaceTaskId}`);
        const pct = result.total > 0 ? Math.round((result.progress / result.total) * 100) : 0;
        const queueBadge = document.getElementById('excel-face-queue-badge');
        const btnCancel = document.getElementById('btn-cancel-excel-face');

        if (result.status === 'queued') {
            if (queueBadge) queueBadge.style.display = 'inline-block';
            if (btnCancel) {
                btnCancel.style.display = 'inline-block';
                btnCancel.disabled = false;
                btnCancel.textContent = 'Hủy chờ';
            }
            document.getElementById('excel-face-progress-title').textContent = 'Đang chờ lượt (có đợt khác đang quét)...';
            document.getElementById('excel-face-progress-percent').textContent = '0%';
            document.getElementById('excel-face-progress-fill').style.width = '0%';
            document.getElementById('excel-face-progress-detail').textContent = 'Hàng đợi tuần tự';
            setTimeout(checkExcelFaceProgress, 1500);
            return;
        }

        if (queueBadge) queueBadge.style.display = 'none';

        if (result.status === 'running') {
            if (btnCancel) {
                btnCancel.style.display = 'inline-block';
                btnCancel.disabled = false;
                btnCancel.textContent = 'Dừng quét';
            }
            document.getElementById('excel-face-progress-title').textContent =
                result.current ? `Đang xuất: ${result.current}` : 'Đang quét mặt...';
            document.getElementById('excel-face-progress-percent').textContent = `${pct}%`;
            document.getElementById('excel-face-progress-fill').style.width = `${pct}%`;
            document.getElementById('excel-face-progress-detail').textContent =
                `${result.progress}/${result.total} file`;
            setTimeout(checkExcelFaceProgress, 1000);
            return;
        }

        if (result.status === 'cancelling') {
            if (btnCancel) {
                btnCancel.disabled = true;
                btnCancel.textContent = 'Đang dừng...';
            }
            document.getElementById('excel-face-progress-title').textContent = 'Đang dừng quét mặt...';
            setTimeout(checkExcelFaceProgress, 1000);
            return;
        }

        if (result.status === 'cancelled') {
            if (btnCancel) btnCancel.style.display = 'none';
            document.getElementById('excel-face-progress-title').textContent = 'Đã dừng quét!';
            document.getElementById('excel-face-progress-detail').textContent = `Đã giữ lại ${result.progress || 0} file Word hoàn tất`;
            showToast(`Đã dừng quét! Đã bảo toàn ${result.progress || 0} file Word cá nhân.`, 'warning');
            loadExcelFaceFiles();
            loadAggregateReports();
            setTimeout(() => {
                document.getElementById('excel-face-progress-section').style.display = 'none';
            }, 3500);
            return;
        }

        if (result.status === 'interrupted') {
            if (btnCancel) btnCancel.style.display = 'none';
            document.getElementById('excel-face-progress-title').textContent = 'Đợt quét bị gián đoạn do máy chủ tắt ngang';
            showToast('Đợt quét bị gián đoạn do máy chủ tắt ngang. Bạn có thể bấm quét lại.', 'warning');
            loadExcelFaceFiles();
            setTimeout(() => {
                document.getElementById('excel-face-progress-section').style.display = 'none';
            }, 4000);
            return;
        }

        if (result.status === 'completed') {
            if (btnCancel) btnCancel.style.display = 'none';
            document.getElementById('excel-face-progress-percent').textContent = '100%';
            document.getElementById('excel-face-progress-fill').style.width = '100%';
            showToast(`Hoàn thành! Đã tạo file Word theo nhân viên và báo cáo giải trình tổng hợp.`, 'success', 'Xem Kết Quả Ngay', () => navigateToResults('summary'));
            loadExcelFaceFiles();
            loadAggregateReports();
            setTimeout(() => {
                document.getElementById('excel-face-progress-section').style.display = 'none';
            }, 3000);
            return;
        }

        if (result.status === 'failed') {
            if (btnCancel) btnCancel.style.display = 'none';
            showToast('Lỗi: ' + (result.errors[0] || 'Không rõ'), 'error');
        }
    } catch (error) {
        console.error('Lỗi kiểm tra tiến độ quét mặt:', error);
        setTimeout(checkExcelFaceProgress, 2000);
    }
}

async function cancelExcelFaceTask() {
    if (!excelFaceTaskId) return;
    const btn = document.getElementById('btn-cancel-excel-face');
    if (btn) {
        btn.disabled = true;
        btn.textContent = 'Đang dừng...';
    }
    try {
        const res = await apiPost(`/api/task/cancel/${excelFaceTaskId}`);
        if (res.success) {
            showToast('Đã gửi yêu cầu dừng quét', 'info');
        } else {
            showToast('Không thể hủy tác vụ: ' + (res.error || ''), 'error');
            if (btn) btn.disabled = false;
        }
    } catch (e) {
        showToast('Lỗi gửi lệnh dừng: ' + e.message, 'error');
        if (btn) btn.disabled = false;
    }
}

async function cancelPdfFaceTask() {
    if (!pdfFaceTaskId) return;
    const btn = document.getElementById('btn-cancel-pdf-face');
    if (btn) {
        btn.disabled = true;
        btn.textContent = 'Đang dừng...';
    }
    try {
        const res = await apiPost(`/api/task/cancel/${pdfFaceTaskId}`);
        if (res.success) {
            showToast('Đã gửi yêu cầu dừng quét PDF', 'info');
        } else {
            showToast('Không thể hủy tác vụ: ' + (res.error || ''), 'error');
            if (btn) btn.disabled = false;
        }
    } catch (e) {
        showToast('Lỗi gửi lệnh dừng: ' + e.message, 'error');
        if (btn) btn.disabled = false;
    }
}

async function confirmClearFaceCache() {
    if (!await confirmAction('Bạn có chắc chắn muốn xóa và làm mới toàn bộ bộ nhớ đệm vector khuôn mặt SQLite không?\n\nLưu ý: Các đợt quét tiếp theo sẽ tính toán lại vector cho ảnh camera và ảnh chân dung.')) {
        return;
    }
    try {
        const res = await apiPost('/api/cache/face/clear');
        if (res.success) {
            showToast('Đã làm mới bộ nhớ đệm khuôn mặt SQLite thành công!', 'success');
        } else {
            showToast('Lỗi: ' + (res.error || 'Không thể xóa cache'), 'error');
        }
    } catch (e) {
        showToast('Lỗi: ' + e.message, 'error');
    }
}

async function loadExcelFaceFiles(btn) {
    return withLoading(btn, async () => {
        try {
            const data = await apiGet('/api/excel/face/files');
            const container = document.getElementById('excel-face-files');
            if (data.folders && data.folders.length > 0) {
                let html = '';
                for (const folder of data.folders) {
                    const safeFolder = folder.folder.replace(/'/g, "\\'");
                    html += `
                        <div class="card" style="margin-bottom: 16px;">
                            <div class="card-header">
                                <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                                    <h3 style="margin:0;">${folder.folder} (${folder.count} file)</h3>
                                    <button class="btn btn-icon-danger btn-sm" title="Xóa đợt kết quả này" onclick="deleteExcelFaceFolder('${safeFolder}')">
                                        ${TRASH_ICON_SVG}<span>Xóa đợt</span>
                                    </button>
                                </div>
                            </div>
                            <div class="card-body">
                                <div class="table-container">
                                    <table class="data-table">
                                        <thead>
                                            <tr>
                                                <th>STT</th>
                                                <th>Tên File Word</th>
                                                <th>Kích Thước</th>
                                                <th>Thao Tác</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            ${folder.files.map((file, idx) => {
                                                const safeFile = file.name.replace(/'/g, "\\'");
                                                return `
                                                    <tr>
                                                        <td>${idx + 1}</td>
                                                        <td>${file.name}</td>
                                                        <td>${formatFileSize(file.size)}</td>
                                                        <td>
                                                            <div style="display:flex;gap:6px;align-items:center;">
                                                                <a href="/api/excel/face/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}"
                                                                   class="btn btn-primary btn-sm">Tải về</a>
                                                                <button class="btn btn-icon-danger btn-sm" title="Xóa file này" onclick="deleteExcelFaceFile('${safeFolder}', '${safeFile}')">
                                                                    ${TRASH_ICON_SVG}
                                                                </button>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                `;
                                            }).join('')}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    `;
                }
                container.innerHTML = html;
            } else {
                container.innerHTML = '<p class="empty-message">Chưa có file Word nào</p>';
            }
        } catch (error) {
            console.error('Lỗi load Excel face files:', error);
        }
    });
}

// ==================== DELETE ACTIONS ====================

async function deleteExcelUpload(filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file "${filename}"?\nHành động này không thể hoàn tác.`)) return;
    try {
        const res = await apiPost('/api/excel/delete-upload', { filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file thành công', 'success');
            loadExcelUploads();
        } else {
            showToast(res.error || 'Lỗi khi xóa file', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteExcelExtractedFolder(folder) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa toàn bộ đợt "${folder}"?\nTất cả file Word trong đợt này sẽ bị xóa.`)) return;
    try {
        const res = await apiPost('/api/excel/delete-folder', { folder });
        if (res.success) {
            showToast(res.message || 'Đã xóa đợt thành công', 'success');
            loadExcelExtractedFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa đợt', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteExcelExtractedFile(folder, filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
    try {
        const res = await apiPost('/api/excel/delete-file', { folder, filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file', 'success');
            loadExcelExtractedFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa file', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteExcelFaceFolder(folder) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa toàn bộ kết quả quét mặt "${folder}"?`)) return;
    try {
        const res = await apiPost('/api/excel/face/delete-folder', { folder });
        if (res.success) {
            showToast(res.message || 'Đã xóa thành công', 'success');
            loadExcelFaceFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteExcelFaceFile(folder, filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
    try {
        const res = await apiPost('/api/excel/face/delete-file', { folder, filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file', 'success');
            loadExcelFaceFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deletePDFUpload(filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file PDF "${filename}"?`)) return;
    try {
        const res = await apiPost('/api/pdf/delete-upload', { filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file PDF thành công', 'success');
            loadPDFUploads();
        } else {
            showToast(res.error || 'Lỗi khi xóa file', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deletePDFExtractedFolder(folder) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa toàn bộ đợt tách PDF "${folder}"?`)) return;
    try {
        const res = await apiPost('/api/pdf/delete-folder', { folder });
        if (res.success) {
            showToast(res.message || 'Đã xóa đợt tách thành công', 'success');
            loadPDFExtractedFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa đợt', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deletePDFExtractedFile(folder, filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
    try {
        const res = await apiPost('/api/pdf/delete-file', { folder, filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file', 'success');
            loadPDFExtractedFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa file', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deletePDFFaceFolder(folder) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa toàn bộ kết quả quét mặt "${folder}"?`)) return;
    try {
        const res = await apiPost('/api/pdf/face/delete-folder', { folder });
        if (res.success) {
            showToast(res.message || 'Đã xóa thành công', 'success');
            loadPDFFaceFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deletePDFFaceFile(folder, filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
    try {
        const res = await apiPost('/api/pdf/face/delete-file', { folder, filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file', 'success');
            loadPDFFaceFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteResultFile(filename) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa file kết quả "${filename}"?`)) return;
    try {
        const res = await apiPost('/api/results/delete', { filename });
        if (res.success) {
            showToast(res.message || 'Đã xóa file', 'success');
            loadResultFiles();
        } else {
            showToast(res.error || 'Lỗi khi xóa file', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

// Drag-and-drop for Excel upload zone
// Drag-and-drop for Excel upload zone
document.addEventListener('DOMContentLoaded', () => {
    const excelZone = document.getElementById('excel-upload-zone');
    if (excelZone) {
        excelZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            excelZone.classList.add('dragover');
        });
        excelZone.addEventListener('dragleave', () => {
            excelZone.classList.remove('dragover');
        });
        excelZone.addEventListener('drop', (e) => {
            e.preventDefault();
            excelZone.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleExcelFile(files[0]);
            }
        });
    }
});


document.addEventListener('DOMContentLoaded', () => {
    initTabFromUrl(false);
    loadProjects().then(() => {
        if (currentActiveTab === 'photos' && currentPhotoSubtab === 'daily') {
            loadDailyPhotosStats();
        }
    });
});

// ==================== QUẢN LÝ ẢNH & DỰ ÁN ====================

// --- Modal Helper Functions ---
function openModal(modalId) {
    const el = document.getElementById(modalId);
    if (el) el.style.display = 'flex';
}

function closeModal(modalId) {
    const el = document.getElementById(modalId);
    if (el) el.style.display = 'none';
}

function openLightbox(src, caption = '') {
    const modal = document.getElementById('modal-lightbox');
    const img = document.getElementById('lightbox-img');
    const cap = document.getElementById('lightbox-caption');
    if (img) img.src = src;
    if (cap) cap.textContent = caption;
    if (modal) modal.style.display = 'flex';
}

// --- Project Management Functions ---
async function loadProjects() {
    try {
        const res = await apiGet('/api/projects');
        if (res.success && Array.isArray(res.projects)) {
            allProjectsList = res.projects;
            
            // Kiểm tra xem dự án hiện tại có trong danh sách không
            const exists = allProjectsList.some(p => p.name === currentProjectName);
            if (!exists && allProjectsList.length > 0) {
                currentProjectName = res.default_project || allProjectsList[0].name;
            } else if (allProjectsList.length === 0) {
                currentProjectName = res.default_project || 'Chung cư Tân Thuận Đông';
            }
            
            updateProjectDropdowns();
            updateProjectMetrics();
        }
    } catch (e) {
        console.error('Lỗi nạp danh sách dự án:', e);
    }
}

function updateProjectDropdowns() {
    const selects = [
        document.getElementById('global-project-select'),
        document.getElementById('excel-project-select'),
        document.getElementById('pdf-project-select')
    ];
    
    selects.forEach(sel => {
        if (!sel) return;
        const previousVal = sel.value;
        sel.innerHTML = '';
        allProjectsList.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p.name;
            opt.textContent = p.display_name || p.name;
            if (p.name === currentProjectName) {
                opt.selected = true;
            }
            sel.appendChild(opt);
        });
        if (previousVal && allProjectsList.some(p => p.name === previousVal)) {
            sel.value = previousVal;
        }
    });
}

function updateProjectMetrics() {
    const proj = allProjectsList.find(p => p.name === currentProjectName);
    const empCountEl = document.getElementById('project-emp-count');
    const camDaysEl = document.getElementById('project-camera-days');
    const camTotalEl = document.getElementById('project-camera-total');
    
    if (empCountEl) empCountEl.textContent = proj ? proj.employee_count : 0;
    if (camDaysEl) camDaysEl.textContent = proj ? proj.camera_days_count : 0;
    if (camTotalEl) camTotalEl.textContent = proj ? proj.camera_total_images : 0;
}

function onProjectChange(newProject) {
    currentProjectName = newProject;
    
    const selects = [
        document.getElementById('global-project-select'),
        document.getElementById('excel-project-select'),
        document.getElementById('pdf-project-select')
    ];
    selects.forEach(s => {
        if (s && s.value !== newProject) s.value = newProject;
    });
    const reportProjectInput = document.getElementById('project-name');
    if (reportProjectInput) reportProjectInput.value = newProject;
    
    updateProjectMetrics();
    
    if (currentPhotoSubtab === 'daily') {
        loadDailyPhotosStats();
    } else {
        loadPortraits();
    }
}

function openCreateProjectModal() {
    const input = document.getElementById('new-project-name');
    if (input) input.value = '';
    openModal('modal-create-project');
    if (input) input.focus();
}

async function submitCreateProject() {
    const input = document.getElementById('new-project-name');
    const name = input ? input.value.trim() : '';
    if (!name) {
        showToast('Vui lòng nhập tên dự án', 'warning');
        return;
    }
    try {
        const res = await apiPost('/api/projects/create', { name });
        if (res.success) {
            showToast(`Đã tạo dự án: ${res.project}`, 'success');
            closeModal('modal-create-project');
            currentProjectName = res.project;
            await loadProjects();
            onProjectChange(res.project);
        } else {
            showToast(res.error || 'Lỗi tạo dự án', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

function openRenameProjectModal() {
    const input = document.getElementById('rename-project-name');
    if (input) input.value = currentProjectName;
    openModal('modal-rename-project');
    if (input) input.focus();
}

async function submitRenameProject() {
    const input = document.getElementById('rename-project-name');
    const newName = input ? input.value.trim() : '';
    if (!newName || newName === currentProjectName) {
        closeModal('modal-rename-project');
        return;
    }
    try {
        const res = await apiPost('/api/projects/rename', {
            old_name: currentProjectName,
            new_name: newName
        });
        if (res.success) {
            showToast(`Đã đổi tên thành: ${res.display_name || res.name}`, 'success');
            closeModal('modal-rename-project');
            currentProjectName = res.name;
            await loadProjects();
            onProjectChange(res.name);
        } else {
            showToast(res.error || 'Lỗi đổi tên dự án', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function confirmDeleteProject() {
    if (!await confirmAction('Lưu trữ dự án này? Giữ nguyên ảnh và lịch sử.')) return;
    try {
        const res = await apiPost('/api/projects/delete', {name: currentProjectName});
        if (!res.success) throw new Error(res.error);
        await loadProjects();
        allEmployeesList = [];
        renderEmployeeCards([]);
        if (allProjectsList.length) onProjectChange(allProjectsList[0].name);
        showToast(res.message, 'success');
    } catch (err) { showToast(err.message, 'error'); }
}

// --- Sub-tab Photo Navigation ---
function switchPhotoSubtab(subtab, updateUrl = true) {
    currentPhotoSubtab = subtab;
    const isDaily = subtab === 'daily';
    const btnDaily = document.getElementById('subtab-photo-btn-daily');
    const btnPortraits = document.getElementById('subtab-photo-btn-portraits');
    const contentDaily = document.getElementById('subtab-photo-daily');
    const contentPortraits = document.getElementById('subtab-photo-portraits');

    if (btnDaily) btnDaily.classList.toggle('active', isDaily);
    if (btnPortraits) btnPortraits.classList.toggle('active', !isDaily);
    if (contentDaily) contentDaily.classList.toggle('active', isDaily);
    if (contentPortraits) contentPortraits.classList.toggle('active', !isDaily);

    if (isDaily) {
        loadDailyPhotosStats();
    } else {
        loadPortraits();
    }

    if (updateUrl && currentActiveTab === 'photos') {
        updateUrlParams('photos', subtab, false);
    }
}

// --- Daily Camera Photos Management ---
function selectedPhotoPeriod() {
    const input = document.getElementById('photo-period');
    if (input && !input.value) {
        const now = new Date();
        const defaultVal = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2, '0');
        if (input._flatpickr) {
            input._flatpickr.setDate(defaultVal, false);
        } else {
            input.value = defaultVal;
        }
    }
    return input ? input.value : '';
}
function selectedPhotoDate(day = activeSelectedDay) {
    return AttendanceDate.fullDate(selectedPhotoPeriod(), day);
}
async function mapLegacyPhotoDay() {
    const folder = prompt('Thư mục ngày cũ cần gán (VD: 01):');
    if (!folder) return;
    const target = prompt('Ngày đầy đủ của TẤT CẢ ảnh trong thư mục (YYYY-MM-DD):');
    if (!target) return;
    if (!await confirmAction('Chỉ xác nhận nếu mọi ảnh trong thư mục thuộc cùng ngày. Nếu lẫn tháng, hãy phân loại từng ảnh trước.')) return;
    const result = await apiPost('/api/photos/daily/legacy-map', {
        project: currentProjectName, folder, date: target, reviewer: 'system', confirm_single_period: true
    });
    showToast(result.success ? 'Đã lưu ngày của thư mục cũ' : result.error, result.success ? 'success' : 'error');
    if (result.success) loadDailyPhotosStats();
}

async function loadDailyPhotosStats(btn) {
    return withLoading(btn, async () => {
        try {
            const res = await apiGet(`/api/photos/daily?project=${encodeURIComponent(currentProjectName)}&period=${selectedPhotoPeriod()}`);
            const grid = document.getElementById('days-picker-grid');
            if (!grid) return;
            grid.innerHTML = '';

            let lastDayWithImages = null;

            if (res.success && Array.isArray(res.days)) {
                res.days.forEach(d => {
                    if (d.has_images) {
                        lastDayWithImages = d.day;
                    }
                    const b = document.createElement('button');
                    b.type = 'button';
                    b.className = `day-picker-btn ${d.has_images ? 'has-images' : ''} ${d.day === activeSelectedDay ? 'active' : ''}`;
                    b.id = `day-btn-${d.day}`;
                    b.onclick = () => selectDay(d.day);
                    b.innerHTML = `
                        <span class="day-picker-num">${d.day}</span>
                        <span class="day-picker-badge">${d.image_count} ảnh</span>
                    `;
                    grid.appendChild(b);
                });
            }

            // Nếu ngày đang chọn không có ảnh mà dự án có ngày khác có ảnh, tự động chuyển sang ngày có ảnh gần nhất
            if (lastDayWithImages) {
                const currentDayObj = res.days && res.days.find(d => d.day === activeSelectedDay);
                if (!currentDayObj || !currentDayObj.has_images) {
                    activeSelectedDay = lastDayWithImages;
                }
            }

            if (res.days && !res.days.some(d => d.day === activeSelectedDay)) activeSelectedDay = '01';
            selectDay(activeSelectedDay);
            updateProjectMetrics();
        } catch (err) {
            console.error('Lỗi tải thống kê ảnh ngày:', err);
        }
    });
}

function selectDay(day) {
    activeSelectedDay = day;
    document.querySelectorAll('.day-picker-btn').forEach(btn => {
        btn.classList.toggle('active', btn.id === `day-btn-${day}`);
    });
    const badgeEl = document.getElementById('selected-day-badge');
    if (badgeEl) badgeEl.textContent = selectedPhotoDate(day);

    loadSelectedDayPhotos(day);
}

async function loadSelectedDayPhotos(day) {
    const gallery = document.getElementById('daily-photos-gallery');
    const subinfo = document.getElementById('selected-day-subinfo');
    const btnDeleteAll = document.getElementById('btn-delete-all-day-photos');
    if (!gallery) return;

    try {
        const res = await apiGet(`/api/photos/daily/${selectedPhotoDate(day)}?project=${encodeURIComponent(currentProjectName)}`);
        if (!res.success) {
            gallery.textContent = res.error || 'Cần kiểm tra ngày ảnh';
            if (btnDeleteAll) btnDeleteAll.style.display = 'none';
            return;
        }
        if (res.success && Array.isArray(res.photos)) {
            const count = res.photos.length;
            if (subinfo) subinfo.textContent = `(${count} ảnh)`;
            if (btnDeleteAll) btnDeleteAll.style.display = count > 0 ? 'inline-flex' : 'none';

            if (count === 0) {
                gallery.innerHTML = `
                    <div style="grid-column: 1 / -1; padding: 24px; text-align: center; color: var(--text-muted); background: var(--bg-subtle); border-radius: var(--radius);">
                        Chưa có ảnh camera nào cho ngày ${day}. Hãy kéo thả hoặc nhấp vào ô phía trên để tải ảnh lên.
                    </div>
                `;
                return;
            }

            const shiftGroups = [
                ['morning', 'Ca sáng · 05:00–15:00'],
                ['afternoon', 'Ca chiều · 16:00–24:00'],
                ['outside', 'Ngoài ca · trước 05:00 hoặc sau 15:00 đến trước 16:00'],
                ['unknown', 'Chưa rõ giờ gửi'],
            ];
            gallery.innerHTML = shiftGroups.map(([shift, label]) => {
                const photos = res.photos.filter(p => (p.shift || 'unknown') === shift);
                if (!photos.length) return '';
                return `<div style="grid-column:1 / -1; font-weight:600; padding-top:12px">${label} (${photos.length} ảnh)</div>` + photos.map(p => `
                <div class="thumb-card">
                    <div class="thumb-img-wrap" onclick="openLightbox('${p.url}', 'Ảnh camera ngày ${day} - ${p.filename}')">
                        <img src="${p.url}" alt="${p.filename}" loading="lazy">
                    </div>
                    <div class="thumb-info">
                        <div class="thumb-name" title="${p.filename}">${p.filename}</div>
                        <div class="thumb-size">${p.size_kb} KB${p.send_time ? ` · ${escapeHtml(p.send_time)}` : ''}</div>
                    </div>
                    <div class="thumb-actions">
                        <button class="btn btn-icon-danger btn-sm" onclick="deleteDailyPhoto('${p.filename}')" title="Xóa ảnh này">
                            ${TRASH_ICON_SVG}
                        </button>
                    </div>
                </div>
            `).join('');
            }).join('');
        }
    } catch (err) {
        console.error('Lỗi nạp ảnh ngày:', err);
    }
}

async function exportDailyPhotosWord(btn) {
    if (btn && btn.disabled) return;
    return withLoading(btn, async () => {
        let objectUrl = null;
        let link = null;
        try {
            const project = currentProjectName;
            const date = selectedPhotoDate();
            const query = new URLSearchParams({project, date});
            const response = await fetch(`/api/photos/daily/export-word?${query}`, {cache: 'no-store'});
            if (!response.ok) {
                const result = await parseJsonResponse(response);
                throw new Error(result.error || 'Không thể xuất Word');
            }
            const contentType = (response.headers.get('Content-Type') || '').split(';')[0].trim().toLowerCase();
            if (contentType !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
                throw new Error('Máy chủ chưa trả về file Word. Vui lòng khởi động lại phần mềm rồi xuất lại.');
            }
            const blob = await response.blob();
            const signature = new Uint8Array(await blob.slice(0, 4).arrayBuffer());
            if (signature.length !== 4 || signature[0] !== 0x50 || signature[1] !== 0x4b || signature[2] !== 0x03 || signature[3] !== 0x04) {
                throw new Error('File Word máy chủ trả về không hợp lệ. Vui lòng xuất lại.');
            }
            const disposition = response.headers.get('Content-Disposition') || '';
            const encodedName = disposition.match(/filename\*=UTF-8''([^;]+)/i);
            const plainName = disposition.match(/filename="([^"]+)"|filename=([^;]+)/i);
            const filename = encodedName ? decodeURIComponent(encodedName[1])
                : plainName ? (plainName[1] || plainName[2]).trim() : `Anh_${project}_${date}.docx`;
            objectUrl = URL.createObjectURL(blob);
            link = document.createElement('a');
            link.href = objectUrl;
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            showToast(`Đã xuất Word ảnh ngày ${date}`, 'success');
        } catch (err) {
            showToast(err.message || 'Không thể xuất Word', 'error');
        } finally {
            if (link) link.remove();
            if (objectUrl) setTimeout(() => URL.revokeObjectURL(objectUrl), 1000);
        }
    });
}

async function handleDailyPhotosUpload(input) {
    if (!input.files || input.files.length === 0) return;
    const formData = new FormData();
    formData.append('project', currentProjectName);
    formData.append('date', selectedPhotoDate());
    for (let i = 0; i < input.files.length; i++) {
        formData.append('files', input.files[i]);
    }

    try {
        showToast(`Đang tải lên ${input.files.length} ảnh camera cho ngày ${activeSelectedDay}...`, 'info');
        const res = await fetch('/api/photos/daily/upload', {
            method: 'POST',
            body: formData
        }).then(r => r.json());

        if (res.success) {
            showToast(`Đã tải lên thành công ${res.saved_count} ảnh cho ngày ${activeSelectedDay}`, 'success');
            input.value = '';
            await loadDailyPhotosStats();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi tải ảnh', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteDailyPhoto(filename) {
    if (!await confirmAction(`Xóa ảnh camera "${filename}" của ngày ${activeSelectedDay}?`)) return;
    try {
        const res = await apiPost('/api/photos/daily/delete', {
            project: currentProjectName,
            date: selectedPhotoDate(),
            filename: filename
        });
        if (res.success) {
            showToast('Đã xóa ảnh', 'success');
            await loadDailyPhotosStats();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi khi xóa', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function confirmDeleteAllDayPhotos() {
    if (!await confirmAction(`Bạn có chắc chắn muốn XÓA TOÀN BỘ ảnh camera của ngày ${activeSelectedDay} không?`)) return;
    try {
        const res = await apiPost('/api/photos/daily/delete', {
            project: currentProjectName,
            date: selectedPhotoDate(),
            delete_all: true
        });
        if (res.success) {
            showToast(res.message || 'Đã xóa toàn bộ ảnh của ngày', 'success');
            await loadDailyPhotosStats();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi khi xóa', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function confirmDeleteAllDaysPhotos() {
    if (!currentProjectName) {
        showToast('Vui lòng chọn dự án trước', 'warning');
        return;
    }
    const period = selectedPhotoPeriod();
    const periodDisplay = period ? (period.includes('-') ? period.split('-').reverse().join('/') : period) : '';
    const modal = document.getElementById('modal-delete-all-days');
    const projectEl = document.getElementById('delete-all-days-project-name');
    const periodLabelEl = document.getElementById('delete-all-days-period-label');

    if (modal && projectEl && periodLabelEl) {
        projectEl.textContent = currentProjectName;
        periodLabelEl.textContent = periodDisplay || '(Tất cả)';
        const radioPeriod = document.querySelector('input[name="delete-all-days-scope"][value="period"]');
        if (radioPeriod) radioPeriod.checked = true;
        openModal('modal-delete-all-days');
    } else {
        const ok = await confirmAction(`Bạn có chắc chắn muốn XÓA TOÀN BỘ ảnh camera của TẤT CẢ CÁC NGÀY trong tháng ${periodDisplay} (Dự án: "${currentProjectName}") không?\n\nLưu ý: Thao tác này sẽ xóa vĩnh viễn và không thể hoàn tác!`);
        if (ok) {
            submitDeleteAllDaysPhotosDirect('period');
        }
    }
}

async function submitDeleteAllDaysPhotos() {
    const scopeRadio = document.querySelector('input[name="delete-all-days-scope"]:checked');
    const scope = scopeRadio ? scopeRadio.value : 'period';
    const btn = document.getElementById('btn-submit-delete-all-days');
    closeModal('modal-delete-all-days');
    await submitDeleteAllDaysPhotosDirect(scope, btn);
}

async function submitDeleteAllDaysPhotosDirect(scope = 'period', btn = null) {
    if (!currentProjectName) return;
    const period = selectedPhotoPeriod();
    const payload = {
        project: currentProjectName,
        delete_all_days: true,
        scope: scope,
        period: period
    };

    try {
        showToast('Đang tiến hành xóa ảnh của các ngày...', 'info');
        const res = await apiPost('/api/photos/daily/delete', payload);
        if (res.success) {
            showToast(res.message || `Đã xóa thành công ${res.deleted_count || 0} ảnh`, 'success');
            await loadDailyPhotosStats();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi khi xóa ảnh', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

// --- Employee Portraits Management & Transfer ---
async function loadPortraits(btn) {
    return withLoading(btn, async () => {
        try {
            const searchInput = document.getElementById('employee-search-input');
            const kw = searchInput ? searchInput.value.trim() : '';
            const res = await apiGet(`/api/portraits?project=${encodeURIComponent(currentProjectName)}&search=${encodeURIComponent(kw)}`);
            if (res.success && Array.isArray(res.employees)) {
                allEmployeesList = res.employees;
                renderEmployeeCards(allEmployeesList);
                const supplementSelect = document.getElementById('supp-employee');
                if (supplementSelect) { supplementSelect.replaceChildren(); for (const employee of allEmployeesList) { const option = document.createElement('option'); option.value = employee.employee_id; option.textContent = employee.name; supplementSelect.appendChild(option); } }
                updateProjectMetrics();
            }
        } catch (err) {
            console.error('Lỗi nạp danh sách nhân viên:', err);
        }
    });
}

function renderEmployeeCards(list) {
    const grid = document.getElementById('employee-cards-grid');
    if (!grid) return;

    if (!list || list.length === 0) {
        grid.innerHTML = `
            <div style="grid-column: 1 / -1; padding: 32px; text-align: center; color: var(--text-muted); background: var(--bg-subtle); border-radius: var(--radius);">
                Không tìm thấy nhân viên nào trong dự án "${currentProjectName}". Nhấp vào <strong>"+ Thêm Nhân Viên"</strong> để tạo mới.
            </div>
        `;
        return;
    }

    grid.innerHTML = list.map(emp => {
        const words = emp.name.trim().split(/\s+/);
        const initials = words.length > 1 
            ? (words[0][0] + words[words.length - 1][0]).toUpperCase()
            : (words[0] ? words[0].substring(0, 2).toUpperCase() : 'NV');
            
        const avatarHtml = emp.avatar_url
            ? `<img src="${emp.avatar_url}" alt="${emp.name}" onclick="openLightbox('${emp.avatar_url}', '${emp.name}')">`
            : `<div class="employee-avatar-placeholder">${initials}</div>`;
        const badgeHtml = emp.image_count > 0
            ? `<span class="employee-badge has-photo">${emp.image_count} ảnh chân dung</span>`
            : `<span class="employee-badge no-photo">Chưa có ảnh</span>`;

        const safeName = emp.employee_id.replace(/'/g, "\\'");
        return `
            <div class="employee-card">
                <button class="employee-card-delete-btn" onclick="deleteEmployee('${safeName}')" title="Xóa nhân viên">
                    ${TRASH_ICON_SVG}
                </button>
                <div class="employee-avatar-wrap">
                    ${avatarHtml}
                </div>
                <div class="employee-name" title="${emp.name}">${emp.name}</div>
                ${badgeHtml}
                <div class="employee-badge" style="background: ${emp.payroll_code ? 'rgba(16, 185, 129, 0.12)' : 'var(--bg-subtle)'}; color: ${emp.payroll_code ? '#059669' : 'var(--text-muted)'}; font-weight: 500;">
                    Mã chấm công: ${escapeHtml(emp.payroll_code || "Chưa có")}
                </div>
                <div class="employee-badge">Mã nội bộ: ${escapeHtml(emp.internal_code || "Chưa tạo")}</div>
                <div class="employee-card-actions">
                    <button class="btn btn-secondary btn-sm" onclick="openEmployeePhotosModal('${safeName}')" title="Xem hoặc thêm ảnh chân dung">
                        Xem / Thêm ảnh
                    </button>
                    <button class="btn btn-secondary btn-sm" onclick="openTransferModal('${safeName}')" title="Thuyên chuyển sang dự án khác">
                        Chuyển
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

function onEmployeeSearch(keyword) {
    const kw = (keyword || '').toLowerCase().trim();
    if (!kw) {
        renderEmployeeCards(allEmployeesList);
        return;
    }
    const filtered = allEmployeesList.filter(emp => emp.name.toLowerCase().includes(kw));
    renderEmployeeCards(filtered);
}

let employeeRosterRows = [];
let employeeRosterProject = '';

async function openCreateEmployeeModal() {
    const project = currentProjectName;
    const nameInput = document.getElementById('new-employee-name');
    const codeInput = document.getElementById('new-employee-code');
    const fileInput = document.getElementById('new-employee-photos');
    if (nameInput) nameInput.value = '';
    if (codeInput) codeInput.value = '';
    if (fileInput) fileInput.value = '';
    employeeRosterRows = [];
    employeeRosterProject = project;
    const select = document.getElementById('new-employee-roster');
    select.replaceChildren(new Option('Đang đọc bảng đối chiếu...', ''));
    nameInput.readOnly = true;
    codeInput.readOnly = true;
    openModal('modal-create-employee');
    try {
        const res = await apiGet(`/api/portraits/audit?project=${encodeURIComponent(project)}`);
        if (project !== currentProjectName) return;
        if (!res.success) throw new Error(res.error);
        employeeRosterRows = res.audit.rows;
        select.replaceChildren(new Option('-- Chọn tên và mã trong bảng --', ''));
        employeeRosterRows.forEach((row, index) => {
            const option = new Option(`${row.payroll_code || 'Thiếu mã'} — ${row.name}`, String(index));
            option.disabled = !row.payroll_code || employeeRosterRows.filter(r => r.payroll_code === row.payroll_code).length !== 1 || res.audit.employees.some(e => e.codes.includes(row.payroll_code));
            select.add(option);
        });
        const loaded = !!res.audit.filename;
        nameInput.readOnly = loaded;
        codeInput.readOnly = loaded;
        document.getElementById('new-employee-roster-hint').textContent = loaded
            ? `Bảng: ${res.audit.filename}. Người đã có mã trong dự án được khóa để tránh tạo trùng. Kiểm tra ảnh đúng người trước khi lưu.`
            : 'Chưa chọn bảng đối chiếu. Nhập tay chưa xác minh được với bảng; nên chọn bảng Excel/PDF trước.';
    } catch (err) {
        document.getElementById('new-employee-roster-hint').textContent = `Không đọc được bảng: ${err.message}. Đóng cửa sổ và thử lại.`;
    }
}

function selectEmployeeRosterRow() {
    const value = document.getElementById('new-employee-roster').value;
    const row = value === '' ? null : employeeRosterRows[Number(value)];
    document.getElementById('new-employee-name').value = row?.name || '';
    document.getElementById('new-employee-code').value = row?.payroll_code || '';
}

async function openEmployeeAudit(project = currentProjectName) {
    openModal('modal-employee-audit');
    const content = document.getElementById('employee-audit-content');
    content.textContent = 'Đang đối chiếu...';
    try {
        const res = await apiGet(`/api/portraits/audit?project=${encodeURIComponent(project)}`);
        if (!res.success) throw new Error(res.error);
        const audit = res.audit;
        const esc = escapeHtml;
        content.innerHTML = `<div class="roster-heading"><strong>${esc(project)}</strong><span>${esc(audit.filename || 'Chưa chọn bảng')}</span></div>
            <div class="roster-controls"><input class="form-input" id="audit-search" placeholder="Tìm tên hoặc mã"><label><input type="checkbox" id="audit-only-issues"> Chỉ cần kiểm tra</label><button class="btn btn-primary btn-sm" id="bulk-import-roster-btn">Nhập từ bảng</button></div>
            <p class="roster-muted">${audit.employees.length} hồ sơ hiện tại · ${audit.missing_profiles.length} dòng chưa có hồ sơ theo mã</p>
            <div class="roster-table-wrap"><table class="data-table"><thead><tr><th>Nhân viên / Mã</th><th>Đối chiếu</th><th>Dự án khác</th><th>Thao tác</th></tr></thead><tbody id="audit-rows"></tbody></table></div>
            <details><summary>Chưa có hồ sơ (${audit.missing_profiles.length})</summary><p>${esc(audit.missing_profiles.map(r => `${r.payroll_code || 'Thiếu mã'} — ${r.name}`).join('; ') || 'Không có')}</p></details>`;
        const render = () => {
            const term = document.getElementById('audit-search').value.toLocaleLowerCase();
            const onlyIssues = document.getElementById('audit-only-issues').checked;
            const rows = audit.employees.filter(e => `${e.name} ${e.codes.join(' ')}`.toLocaleLowerCase().includes(term) && (!onlyIssues || e.issues.length));
            content.querySelector('#audit-rows').innerHTML = rows.map(e => {
                const current = e.other_projects.filter(p => p.current);
                const history = e.other_projects.filter(p => !p.current);
                const matched = e.roster_rows.length === 1;
                return `<tr><td><strong>${esc(e.name)}</strong><div class="roster-muted">${esc(e.codes.join(', ') || 'Chưa có mã')}</div></td>
                    <td><span class="roster-badge ${e.issues.length ? 'needs-review' : ''}">${e.issues.length ? 'Cần kiểm tra' : 'Khớp bảng'}</span>${e.issues.length ? `<details><summary>Chi tiết</summary>${e.issues.map(i => `<div>${esc(i)}</div>`).join('')}</details>` : ''}</td>
                    <td>${current.map(p => `<div>${esc(p.project)} <span class="roster-muted">· Hiện tại</span></div>`).join('') || '—'}${history.length ? `<details><summary>Lịch sử (${history.length})</summary>${history.map(p => `<div>${esc(p.project)} · Đã kết thúc</div>`).join('')}</details>` : ''}</td>
                    <td><div class="roster-row-actions"><button class="btn btn-secondary btn-sm" data-audit-photos="${esc(e.employee_id)}">Ảnh</button>${matched ? `<button class="btn btn-secondary btn-sm" data-exclusive="${esc(e.employee_id)}">Chỉ thuộc dự án này</button>` : ''}</div></td></tr>`;
            }).join('') || '<tr><td colspan="4">Không có hồ sơ phù hợp</td></tr>';
            content.querySelectorAll('[data-audit-photos]').forEach(btn => btn.onclick = () => { if (currentProjectName !== project) return showToast('Chọn dự án này trong quản lý ảnh để xem ảnh', 'info'); closeModal('modal-employee-audit'); openEmployeePhotosModal(btn.dataset.auditPhotos); });
            content.querySelectorAll('[data-exclusive]').forEach(btn => btn.onclick = () => confirmExclusiveProject(project, btn.dataset.exclusive));
        };
        content.querySelector('#audit-search').oninput = render;
        content.querySelector('#audit-only-issues').onchange = render;
        content.querySelector('#bulk-import-roster-btn').onclick = () => importAllRosterEmployees(project);
        render();
    } catch (err) { content.textContent = `Không kiểm tra được: ${err.message}`; }
}

async function confirmExclusiveProject(project, employeeId) {
    const projectId = allProjectsList.find(p => p.name === project)?.project_id;
    try {
        const preview = await apiPost('/api/portraits/exclusive-project', {project_id: projectId, employee_id: employeeId});
        if (!preview.success) throw new Error(preview.error);
        if (!preview.affected.length) return showToast('Hồ sơ chỉ hoạt động tại dự án này', 'info');
        if (!await confirmAction(`Giữ ${preview.name} (${preview.payroll_code}) tại ${project}.\nLưu trữ các phân công/hồ sơ sau:\n${preview.affected.map(e => `${e.project}: ${e.name}`).join('\n')}\nChỉ xác nhận nếu đây là cùng một người. Ảnh và lịch sử được giữ.`)) return;
        const res = await apiPost('/api/portraits/exclusive-project', {project_id: projectId, employee_id: employeeId, apply: true, confirmation_keys: preview.confirmation_keys});
        if (!res.success) throw new Error(res.error);
        showToast('Đã giữ hồ sơ hoạt động tại dự án được chọn', 'success');
        await loadProjects(); await loadPortraits(); await openEmployeeAudit(project);
    } catch (err) { showToast('Lỗi: ' + err.message, 'error'); }
}

async function importAllRosterEmployees(projectName) {
    const projectId = allProjectsList.find(p => p.name === projectName)?.project_id;
    if (!projectId) return showToast('Vui lòng chọn dự án', 'warning');
    const button = document.getElementById('bulk-import-roster-btn');
    if (button) button.disabled = true;
    try {
        const preview = await apiPost('/api/portraits/import-roster', {project_id: projectId});
        if (!preview.success) throw new Error(preview.error);
        const c = preview.counts;
        await reviewRosterTransfers(projectName, preview);
        const conflicts = preview.rows.filter(r => r.action === 'conflict');
        const message = `Dự án: ${projectName}\nBảng: ${preview.filename}\nTổng: ${preview.total} nhân viên\nThêm mới: ${c.create}; gắn mã hồ sơ cũ: ${c.assign}; khôi phục: ${c.restore}; đã có: ${c.existing}; cần kiểm tra: ${c.conflict}; chờ chuyển dự án: ${c.transfer || 0}.\n` +
            conflicts.map(r => `${r.payroll_code || 'Thiếu mã'} — ${r.name}: ${r.reason}`).join('\n');
        if (!c.create && !c.assign && !c.restore) {
            await confirmAction(message + '\nKhông có hồ sơ nào đủ điều kiện để thêm.');
            return;
        }
        if (!await confirmAction(message + '\nNhập các hồ sơ đủ điều kiện vào dự án này? Các dòng cần kiểm tra sẽ được giữ lại để đối chiếu.')) return;
        const result = await apiPost('/api/portraits/import-roster', {project_id: projectId, apply: true});
        if (!result.success) throw new Error(result.error);
        const n = result.counts;
        showToast(`Đã thêm ${n.create}, gắn mã ${n.assign}, khôi phục ${n.restore}. Đã có ${n.existing}; cần kiểm tra ${n.conflict}.`, n.conflict ? 'warning' : 'success');
        await loadProjects();
        await loadPortraits();
        await openEmployeeAudit(projectName);
        const refreshed = await apiPost('/api/portraits/import-roster', {project_id: projectId});
        if (refreshed.success) await reviewRosterTransfers(projectName, refreshed);
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    } finally {
        if (button) button.disabled = false;
    }
}

async function submitCreateEmployee() {
    if (employeeRosterProject !== currentProjectName) {
        showToast('Dự án đã thay đổi; mở lại cửa sổ thêm nhân viên', 'warning');
        return;
    }
    const nameInput = document.getElementById('new-employee-name');
    const codeInput = document.getElementById('new-employee-code');
    const fileInput = document.getElementById('new-employee-photos');
    const name = nameInput ? nameInput.value.trim() : '';
    const payroll_code = codeInput ? codeInput.value.trim() : '';

    if (!name) {
        showToast('Vui lòng nhập họ tên nhân viên', 'warning');
        return;
    }

    const valid_from = '2000-01-01';
    const reviewer = 'system';

    try {
        const createRes = await apiPost('/api/portraits/employee/create', {
            project_id: allProjectsList.find(p => p.name === currentProjectName)?.project_id,
            project: currentProjectName,
            name: name, payroll_code, valid_from, reviewer
        });

        if (!createRes.success) {
            showToast(createRes.error || 'Lỗi thêm nhân viên', 'error');
            return;
        }

        // Upload photos if any
        if (fileInput && fileInput.files && fileInput.files.length > 0) {
            const formData = new FormData();
            formData.append('project', currentProjectName);
            formData.append('project_id', allProjectsList.find(p => p.name === currentProjectName)?.project_id || '');
            formData.append('employee_id', createRes.employee_id);
            formData.append('reviewer', reviewer);
            for (let i = 0; i < fileInput.files.length; i++) {
                formData.append('files', fileInput.files[i]);
            }
            const uploaded = await fetch('/api/portraits/employee/upload', {
                method: 'POST',
                body: formData
            }).then(r => r.json());
            if (!uploaded.success) {
                closeModal('modal-create-employee');
                await loadPortraits();
                throw new Error('Đã tạo nhân viên nhưng tải ảnh thất bại: ' + uploaded.error);
            }
        }

        showToast(`Đã thêm nhân viên: ${name}`, 'success');
        closeModal('modal-create-employee');
        await loadPortraits();
        await loadProjects();
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

function triggerImportEmployeesFile() {
    const fileInput = document.getElementById('import-employees-file-input');
    if (fileInput) {
        fileInput.value = '';
        fileInput.click();
    }
}

async function handleImportEmployeesFile(input) {
    const file = input?.files?.[0];
    if (!file) return;
    await syncUploadedRoster(file.name, file.name.toLowerCase().endsWith('.pdf') ? 'pdf' : 'excel');
    uploadedRosterSync.localFile = file;
    input.value = '';
}

async function checkRosterBeforeScan(folder, kind, projectName) {
    try {
        const [files, batches] = await Promise.all([apiGet(`/api/${kind}/uploads`), apiGet(`/api/${kind}/files`)]);
        const sourceName = (batches.folders || []).find(b => b.folder === folder)?.source_filename;
        const matches = (files.files || []).filter(f => sourceName ? f.name === sourceName : f.name.replace(/\.[^.]+$/, '') === folder);
        if (matches.length !== 1) {
            showToast('Không tìm thấy file nguồn để đối chiếu; tiếp tục quét dữ liệu đã tách.', 'info');
            return true;
        }
        const projectId = allProjectsList.find(p => p.name === projectName)?.project_id;
        if (!projectId) throw new Error('Chọn dự án đích trước khi quét');
        const imported = await apiPost('/api/portraits/import-file', {filename: matches[0].name, project_id: projectId});
        if (!imported.success) throw new Error(imported.error);
        const preview = await apiPost('/api/portraits/import-roster', {project_id: projectId, authoritative: true});
        if (!preview.success) throw new Error(preview.error);
        if (preview.sync_approved) return true;
        if (preview.counts.transfer || preview.counts.create || preview.counts.assign || preview.counts.restore || preview.counts.conflict || preview.renamed || preview.outside_roster?.length) {
            showToast('Bảng còn thay đổi nhân viên chưa đồng bộ. Quét vẫn tiếp tục với hồ sơ hiện có.', 'info', 'Đồng bộ nhân viên', () => syncUploadedRoster(matches[0].name, kind));
        }
        return true;
    } catch (err) {
        showToast('Không đối chiếu được danh sách: ' + err.message + '. Quét vẫn tiếp tục với hồ sơ hiện có.', 'warning');
        return true;
    }
}

let uploadedRosterSync = null;
async function syncUploadedRoster(filename, kind) {
    uploadedRosterSync = {filename, kind, preview: null, selections: {}, skipped: []};
    document.getElementById('roster-sync-filename').textContent = filename;
    const now = new Date();
    document.getElementById('roster-sync-effective').value = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
    const select = document.getElementById('roster-sync-project');
    select.replaceChildren(new Option('-- Chọn dự án nhận nhân viên --', ''));
    allProjectsList.forEach(p => select.add(new Option(p.name, p.project_id)));
    select.onchange = () => { uploadedRosterSync.selections = {}; uploadedRosterSync.skipped = []; uploadedRosterSync.preview = null; document.getElementById('roster-sync-preview').textContent = 'Bấm Xem trước để đối chiếu với dự án này.'; document.getElementById('roster-sync-apply').disabled = true; };
    document.getElementById('roster-sync-preview').textContent = 'Chọn dự án để xem trước. Hồ sơ nhân viên chưa được thay đổi.';
    document.getElementById('roster-sync-apply').disabled = true;
    document.getElementById('roster-sync-effective').onchange = () => { uploadedRosterSync.preview = null; document.getElementById('roster-sync-apply').disabled = true; };
    openModal('modal-roster-sync');
}

async function previewUploadedRoster() {
    const state = uploadedRosterSync;
    const projectId = document.getElementById('roster-sync-project').value;
    if (!projectId) return showToast('Chọn dự án nhận nhân viên', 'warning');
    const content = document.getElementById('roster-sync-preview');
    state.preview = null;
    document.getElementById('roster-sync-apply').disabled = true;
    content.textContent = 'Đang đọc bảng...';
    try {
        let imported;
        if (state.localFile) {
            const form = new FormData(); form.append('file', state.localFile); form.append('project_id', projectId);
            imported = await fetch('/api/portraits/import-file', {method:'POST', body:form}).then(r => r.json());
        } else {
            imported = await apiPost('/api/portraits/import-file', {filename: state.filename, project_id: projectId});
        }
        if (!imported.success) throw new Error(imported.error);
        const preview = await apiPost('/api/portraits/import-roster', {project_id: projectId, authoritative: true, selections: state.selections, skipped: state.skipped, effective_date: document.getElementById('roster-sync-effective').value});
        if (!preview.success) throw new Error(preview.error);
        if (state !== uploadedRosterSync || projectId !== document.getElementById('roster-sync-project').value) return;
        state.preview = preview;
        const c = preview.counts;
        const labels = {create:'Bổ sung mới', assign:'Cập nhật mã', restore:'Khôi phục', existing:'Đã có', conflict:'Cần chọn hồ sơ', transfer:'Chuyển về dự án', skip:'Không đồng bộ'};
        content.innerHTML = `<div class="roster-summary">${Object.entries(labels).map(([k,v]) => `<span class="roster-badge">${v}: ${c[k] || 0}</span>`).join('')}</div><div class="roster-table-wrap"><table class="data-table"><thead><tr><th>Mã trong bảng</th><th>Nhân viên</th><th>Kết quả</th><th>Đồng bộ</th></tr></thead><tbody>${[...preview.rows].sort((a,b) => Number(b.action === 'conflict')-Number(a.action === 'conflict')).map(r => `<tr><td>${escapeHtml(r.payroll_code || '—')}</td><td>${escapeHtml(r.name)}</td><td>${r.action === 'conflict' && !r.candidates?.length ? 'Kiểm tra dữ liệu' : labels[r.action]}${r.rename ? `<div class="roster-muted">Đổi tên: ${escapeHtml(r.previous_name)} → ${escapeHtml(r.name)}</div>` : ''}${r.source_projects?.length ? `<div class="roster-muted">Từ ${escapeHtml(r.source_projects.join(', '))}</div>` : ''}${r.action === 'conflict' ? `<div class="roster-muted">${escapeHtml(r.reason)}</div>${r.candidates?.length ? `<select class="form-input" data-roster-code="${escapeHtml(r.payroll_code)}"><option value="">-- Chọn hồ sơ đúng người --</option>${r.candidates.map(e => `<option value="${escapeHtml(e.employee_id)}">${escapeHtml(e.name)} · ${escapeHtml(e.projects.join(', '))}</option>`).join('')}</select>` : ''}` : ''}</td><td><input type="checkbox" data-sync-row="${escapeHtml(r.row_key)}" ${r.action === 'skip' ? '' : 'checked'} aria-label="Đồng bộ ${escapeHtml(r.name)}"><span class="roster-muted"> ${r.action === 'skip' ? 'Không đồng bộ' : 'Đồng bộ'}</span></td></tr>`).join('')}</tbody></table></div>`;
        content.querySelectorAll('[data-sync-row]').forEach(input => input.onchange = async () => {
            const key = input.dataset.syncRow;
            state.skipped = input.checked ? state.skipped.filter(k => k !== key) : [...new Set([...state.skipped, key])];
            await previewUploadedRoster();
        });
        content.insertAdjacentHTML('beforeend', `<p class="roster-muted">${preview.outside_roster.length} hồ sơ ngoài bảng sẽ được lưu trữ: ${escapeHtml(preview.outside_roster.map(e => e.name).join(', ') || 'Không có')}. Ảnh và lịch sử được giữ.</p>`);
        content.querySelectorAll('[data-roster-code]').forEach(select => select.onchange = async () => {
            if (!select.value) return;
            state.selections[select.dataset.rosterCode] = select.value;
            await previewUploadedRoster();
        });
        if (c.conflict) content.insertAdjacentHTML('afterbegin', `<p class="roster-sync-blocked">Còn ${c.conflict} dòng cần xử lý bên dưới. Chọn hồ sơ nếu có danh sách, hoặc kiểm tra lỗi cụ thể trong bảng.</p>`);
        document.getElementById('roster-sync-apply').disabled = !!c.conflict;

    } catch (err) { content.textContent = err.message; }
}

async function applyUploadedRoster() {
    const state = uploadedRosterSync;
    if (!state?.preview) return;
    const button = document.getElementById('roster-sync-apply'); button.disabled = true;
    try {
        const res = await apiPost('/api/portraits/import-roster', {project_id: state.preview.project_id, apply: true, authoritative: true, roster_token: state.preview.roster_token, effective_date: state.preview.effective_date, selections: state.selections, skipped: state.skipped});
        if (!res.success) throw new Error(res.error);
        showToast(`Đã đồng bộ: bổ sung ${res.counts.create + res.counts.restore}, chuyển ${res.counts.transfer}, cập nhật mã ${res.counts.assign}, đổi tên ${res.renamed || 0}`, 'success');
        await loadProjects(); await loadPortraits(); closeModal('modal-roster-sync');
        showToast('Đã lưu quyết định đồng bộ. Bạn có thể bấm Quét mặt.', 'success');
    } catch (err) { button.disabled = false; showToast(err.message, 'error'); }
}

async function reviewRosterTransfers(projectName, preview, targetContent = null) {
    const rows = preview.rows.filter(r => r.action === 'transfer');
    if (!rows.length) return;
    const content = targetContent || document.getElementById('employee-audit-content');
    content.querySelector('#roster-transfer-review')?.remove();
    const panel = document.createElement('div');
    panel.id = 'roster-transfer-review';
    const today = new Date();
    const dateValue = `${today.getFullYear()}-${String(today.getMonth()+1).padStart(2,'0')}-${String(today.getDate()).padStart(2,'0')}`;
    panel.innerHTML = `<h4>Nhân viên có hồ sơ ở dự án khác</h4><p>Chọn đúng hồ sơ và ngày chuyển. Dự án đích dùng mã trong bảng; dự án nguồn giữ lịch sử đến ngày chuyển.</p>`;
    for (const row of rows) {
        const item = document.createElement('div');
        item.style.cssText = 'padding:12px 0;border-bottom:1px solid var(--border-color);display:flex;gap:8px;flex-wrap:wrap;align-items:center';
        const label = document.createElement('strong');
        label.textContent = `${row.name} — Mã đích: ${row.payroll_code}`;
        const select = document.createElement('select');
        select.className = 'form-input';
        select.add(new Option('-- Chọn hồ sơ nguồn --', ''));
        row.transfer_candidates.forEach((c,i) => select.add(new Option(`${c.source_project}: ${c.name} (mã ${c.source_codes.join(', ') || 'chưa có'})`, String(i))));
        const date = document.createElement('input');
        date.type = 'date'; date.value = dateValue; date.className = 'form-input'; date.setAttribute('aria-label', 'Ngày chuyển dự án');
        const button = document.createElement('button');
        button.className = 'btn btn-primary btn-sm'; button.textContent = 'Xác nhận chuyển';
        button.onclick = async () => {
            const candidate = select.value === '' ? null : row.transfer_candidates[Number(select.value)];
            if (!candidate || !date.value) return showToast('Chọn hồ sơ nguồn và ngày chuyển', 'warning');
            if (!await confirmAction(`Chuyển ${candidate.name} từ ${candidate.source_project} sang ${projectName} ngày ${date.value}, dùng mã ${row.payroll_code} từ bảng? Xác nhận ảnh là đúng người trước khi chuyển.`)) return;
            button.disabled = true;
            try {
                const res = await apiPost('/api/portraits/roster-transfer', {project_id: preview.project_id, source_project_id: candidate.source_project_id, employee_id: candidate.employee_id, payroll_code: row.payroll_code, effective_date: date.value});
                if (!res.success) throw new Error(res.error);
                button.textContent = 'Đã chuyển'; select.disabled = true; date.disabled = true;
                showToast(res.message, 'success'); await loadProjects(); await loadPortraits();
            } catch (err) { button.disabled = false; showToast('Lỗi: ' + err.message, 'error'); }
        };
        item.append(label, select, date, button); panel.append(item);
    }
    content.prepend(panel);
    panel.scrollIntoView({block: 'nearest'});
}

async function syncEmployeesFromExcel() {
    const filename = document.getElementById('excel-filename')?.textContent.trim();
    if (!filename) return showToast('Chọn file trước khi đồng bộ', 'warning');
    await syncUploadedRoster(filename, 'excel');
}

async function syncEmployeesFromPDF() {
    const filename = document.getElementById('pdf-filename')?.textContent.trim();
    if (!filename) return showToast('Chọn file trước khi đồng bộ', 'warning');
    await syncUploadedRoster(filename, 'pdf');
}

function openTransferModal(empName) {
    const nameEl = document.getElementById('transfer-employee-name');
    const selectEl = document.getElementById('transfer-target-project');
    if (nameEl) { nameEl.textContent = allEmployeesList.find(e => e.employee_id === empName)?.name || empName; nameEl.dataset.employeeId = empName; }
    if (selectEl) {
        selectEl.innerHTML = '';
        const targetProjects = allProjectsList.filter(p => p.name !== currentProjectName);
        if (targetProjects.length === 0) {
            showToast('Chưa có dự án khác để chuyển. Hãy bấm "+ Thêm Dự Án" trước.', 'warning');
            return;
        }
        targetProjects.forEach(p => {
            const opt = document.createElement('option');
            opt.value = p.project_id;
            opt.textContent = p.display_name || p.name;
            selectEl.appendChild(opt);
        });
    }
    openModal('modal-transfer-employee');
}

async function submitTransferEmployee() {
    const nameEl = document.getElementById('transfer-employee-name');
    const selectEl = document.getElementById('transfer-target-project');
    const empName = nameEl ? nameEl.dataset.employeeId : '';
    const targetProj = selectEl ? selectEl.value : '';

    if (!empName || !targetProj) {
        showToast('Vui lòng chọn dự án đích', 'warning');
        return;
    }

    const employee = allEmployeesList.find(item => item.employee_id === empName);
    const effective_date = new Date().toLocaleDateString('en-CA');
    const payroll_code = employee?.payroll_code || '';
    const reviewer = 'system';
    try {
        const res = await apiPost('/api/portraits/employee/transfer', {
            source_project_id: allProjectsList.find(p => p.name === currentProjectName)?.project_id,
            project: currentProjectName,
            target_project_id: targetProj,
            employee_id: empName, effective_date, payroll_code, reviewer
        });

        if (res.success) {
            showToast(res.message || `Đã thuyên chuyển ${empName} sang ${targetProj}`, 'success');
            closeModal('modal-transfer-employee');
            await loadPortraits();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi thuyên chuyển nhân viên', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

function openEmployeePhotosModal(empName) {
    activeEmpModalName = empName;
    const nameEl = document.getElementById('emp-photos-modal-name');
    if (nameEl) { nameEl.textContent = allEmployeesList.find(e => e.employee_id === empName)?.name || empName; nameEl.dataset.employeeId = empName; }
    refreshEmployeePhotosModal();
    if (typeof setupEmployeePhotosModalInteractions === 'function') {
        setupEmployeePhotosModalInteractions();
    }
    openModal('modal-employee-photos');
}

function refreshEmployeePhotosModal() {
    const gallery = document.getElementById('emp-photos-gallery');
    if (!gallery) return;

    const emp = allEmployeesList.find(e => e.employee_id === activeEmpModalName);
    if (!emp || !emp.images || emp.images.length === 0) {
        gallery.innerHTML = `
            <div style="grid-column: 1 / -1; padding: 18px; text-align: center; color: var(--text-muted); background: var(--bg-subtle); border-radius: var(--radius);">
                Chưa có ảnh chân dung nào cho nhân viên này. Nhấp ô phía trên, kéo thả ảnh hoặc nhấn <strong>Ctrl+V</strong> để dán ảnh Zalo.
            </div>
        `;
        return;
    }

    gallery.innerHTML = emp.images.map(img => {
        const url = emp.image_urls[img] || '';
        const encImg = encodeURIComponent(img);
        const displayName = (img.split(/[\/\\]/).pop() || img).replace(/"/g, '&quot;');
        const safeCaption = ((emp.name || activeEmpModalName) + ' - ' + displayName).replace(/'/g, "\\'").replace(/"/g, '&quot;');
        return `
            <div class="thumb-card">
                <div class="thumb-img-wrap" onclick="openLightbox('${url}', '${safeCaption}')">
                    <img src="${url}" alt="${displayName}" loading="lazy">
                </div>
                <div class="thumb-info">
                    <div class="thumb-name" title="${displayName}">${displayName}</div>
                </div>
                <div class="thumb-actions">
                    <button class="btn btn-icon-danger btn-sm" onclick="deleteEmployeePhoto(decodeURIComponent('${encImg}'))" title="Xóa ảnh này">
                        ${TRASH_ICON_SVG}
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

async function handleUploadEmployeePhoto(input) {
    const rawFiles = input && (input.files || (Array.isArray(input) ? input : [input]));
    const files = Array.from(rawFiles || []).filter(f => f && (f.name || f instanceof Blob));
    if (!files || files.length === 0 || !activeEmpModalName) return;
    const formData = new FormData();
    formData.append('project', currentProjectName);
    formData.append('project_id', (typeof allProjectsList !== 'undefined' && Array.isArray(allProjectsList) ? allProjectsList.find(p => p.name === currentProjectName)?.project_id : '') || '');
    formData.append('employee_id', activeEmpModalName);
    formData.append('reviewer', 'system');
    for (let i = 0; i < files.length; i++) {
        formData.append('files', files[i]);
    }

    try {
        const res = await fetch('/api/portraits/employee/upload', {
            method: 'POST',
            body: formData
        }).then(r => r.json());

        if (res.success) {
            showToast(`Đã thêm ${res.saved_count || files.length} ảnh chân dung cho ${activeEmpModalName}`, 'success');
            if (input && 'value' in input) {
                try { input.value = ''; } catch (_) {}
            }
            const fileInput = typeof document !== 'undefined' && document.getElementById ? document.getElementById('emp-photos-upload-input') : null;
            if (fileInput) fileInput.value = '';
            await loadPortraits();
            refreshEmployeePhotosModal();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi tải ảnh', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteEmployeePhoto(filename) {
    if (!activeEmpModalName) return;
    let cleanFilename = filename;
    try {
        if (cleanFilename && typeof cleanFilename === 'string' && cleanFilename.includes('%')) {
            cleanFilename = decodeURIComponent(cleanFilename);
        }
    } catch (_) {}
    const displayName = cleanFilename ? cleanFilename.split(/[\/\\]/).pop() : '';
    if (!await confirmAction(`Bạn có chắc muốn xóa vĩnh viễn ảnh chân dung "${displayName || cleanFilename}" của ${activeEmpModalName}?`)) return;
    try {
        const res = await apiPost('/api/portraits/employee/delete-photo', {
            project_id: (typeof allProjectsList !== 'undefined' && Array.isArray(allProjectsList) ? allProjectsList.find(p => p.name === currentProjectName)?.project_id : '') || '',
            project: currentProjectName,
            employee_id: activeEmpModalName,
            filename: cleanFilename
        });
        if (res.success) {
            showToast('Đã xóa vĩnh viễn ảnh chân dung', 'success');
            await loadPortraits();
            refreshEmployeePhotosModal();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi khi xóa ảnh', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

async function deleteAllProjectEmployees(button) {
    const projectName = currentProjectName;
    const projectId = allProjectsList.find(p => p.name === projectName)?.project_id;
    if (!projectId) {
        showToast('Vui lòng chọn dự án', 'warning');
        return;
    }
    if (!await confirmAction(`Xóa tất cả nhân viên khỏi danh sách hoạt động của dự án "${projectName}"? Ảnh và lịch sử vẫn được giữ. Các dự án khác không bị ảnh hưởng.`)) return;
    button.disabled = true;
    try {
        const res = await apiPost('/api/portraits/employees/delete-all', {project_id: projectId});
        if (!res.success) throw new Error(res.error || 'Lỗi xóa nhân viên');
        showToast(res.message, 'success');
        await loadProjects();
        await loadPortraits();
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    } finally {
        button.disabled = false;
    }
}

async function deleteEmployee(empName) {
    if (!await confirmAction(`Bạn có chắc chắn muốn xóa nhân viên "${empName}" khỏi danh sách hoạt động? Ảnh và lịch sử vẫn được giữ.`)) return;
    try {
        const res = await apiPost('/api/portraits/employee/delete', {
            project_id: allProjectsList.find(p => p.name === currentProjectName)?.project_id,
            project: currentProjectName,
            employee_id: empName
        });
        if (res.success) {
            showToast(`Đã lưu trữ nhân viên ${empName}`, 'success');
            await loadPortraits();
            await loadProjects();
        } else {
            showToast(res.error || 'Lỗi khi xóa nhân viên', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

// --- Quản lý Thư Mục Ảnh Chân Dung, Kéo Thả Nhiều Ảnh & Dán Ảnh Zalo (Ctrl+V) ---
async function openCurrentEmployeeFolder() {
    if (!activeEmpModalName) {
        showToast('Chưa chọn nhân viên nào', 'warning');
        return;
    }
    try {
        const res = await apiPost('/api/portraits/open-folder', {
            project: currentProjectName,
            employee_id: activeEmpModalName
        });
        if (res.success) {
            showToast(res.message || 'Đã mở thư mục ảnh nhân viên', 'success');
        } else {
            showToast(res.error || 'Không thể mở thư mục ảnh', 'error');
        }
    } catch (err) {
        showToast('Lỗi mở thư mục: ' + err.message, 'error');
    }
}

async function openProjectPortraitsFolder() {
    if (!currentProjectName) {
        showToast('Chưa chọn dự án nào', 'warning');
        return;
    }
    try {
        const res = await apiPost('/api/portraits/open-folder', {
            project: currentProjectName
        });
        if (res.success) {
            showToast(res.message || 'Đã mở thư mục chân dung dự án', 'success');
        } else {
            showToast(res.error || 'Không thể mở thư mục chân dung', 'error');
        }
    } catch (err) {
        showToast('Lỗi mở thư mục: ' + err.message, 'error');
    }
}

function extractImagesFromClipboard(clipboardData) {
    if (!clipboardData) return [];
    const files = [];

    // 1. Check clipboardData.files (file copy from Explorer, Zalo desktop file drop/copy)
    if (clipboardData.files && clipboardData.files.length > 0) {
        for (let i = 0; i < clipboardData.files.length; i++) {
            const f = clipboardData.files[i];
            if (f && ((f.type && f.type.startsWith('image/')) || /\.(jpe?g|png|webp|bmp|gif)$/i.test(f.name))) {
                files.push(f);
            }
        }
    }

    // 2. Check clipboardData.items (image bitmap copy from Zalo, snipping tool, browser)
    if (files.length === 0 && clipboardData.items && clipboardData.items.length > 0) {
        for (let i = 0; i < clipboardData.items.length; i++) {
            const item = clipboardData.items[i];
            if (item && item.type && item.type.startsWith('image/')) {
                const blob = item.getAsFile();
                if (blob) {
                    const ext = (item.type.split('/')[1] || 'png').replace('jpeg', 'jpg');
                    const namedFile = new File(
                        [blob],
                        `zalo_paste_${Date.now()}_${i + 1}.${ext}`,
                        { type: blob.type || item.type }
                    );
                    files.push(namedFile);
                }
            }
        }
    }
    return files;
}

function setupEmployeePhotosModalInteractions() {
    const dropzone = document.getElementById('emp-photos-dropzone');
    if (dropzone && dropzone.dataset.dndInitialized !== 'true') {
        dropzone.dataset.dndInitialized = 'true';

        ['dragenter', 'dragover'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.add('dragover');
            });
        });

        ['dragleave', 'dragend'].forEach(eventName => {
            dropzone.addEventListener(eventName, (e) => {
                e.preventDefault();
                e.stopPropagation();
                dropzone.classList.remove('dragover');
            });
        });

        dropzone.addEventListener('drop', async (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropzone.classList.remove('dragover');

            const dt = e.dataTransfer;
            if (!dt || !dt.files || dt.files.length === 0) return;

            const files = Array.from(dt.files).filter(f =>
                (f.type && f.type.startsWith('image/')) || /\.(jpe?g|png|webp|bmp)$/i.test(f.name)
            );

            if (files.length === 0) {
                showToast('Vui lòng chỉ kéo thả tệp hình ảnh (.jpg, .png, .bmp)', 'warning');
                return;
            }

            await handleUploadEmployeePhoto({ files: files });
        });
    }

    if (typeof window !== 'undefined' && !window._empPhotosPasteBound) {
        window._empPhotosPasteBound = true;
        window.addEventListener('paste', async (e) => {
            const modal = document.getElementById('modal-employee-photos');
            if (!modal || modal.style.display === 'none' || !activeEmpModalName) return;

            const imageFiles = extractImagesFromClipboard(e.clipboardData);
            if (imageFiles && imageFiles.length > 0) {
                e.preventDefault();
                e.stopPropagation();
                showToast(`Đang tải lên ${imageFiles.length} ảnh dán từ Zalo...`, 'info');
                await handleUploadEmployeePhoto({ files: imageFiles });
            }
        });
    }
}

// ==================== TAB: TẢI ẢNH ZALO ====================

let zaloPollTimer = null;
let zaloEventSource = null;
let zaloGroupsList = [];
let zaloIsDownloading = false;
let zaloCurrentQrSource = null;

function initZaloTab() {
    setZaloDatePreset('today');
    loadProjectsForZalo();
    checkZaloStatus();
    checkActiveZaloDownload();

    // Re-render cached groups & restore selected group UI when switching tabs
    if (zaloGroupsList.length > 0) {
        renderZaloGroupItems(zaloGroupsList);
        restoreZaloGroupSelection();
    }
}

// Kiểm tra xem hiện tại có tiến trình tải nào đang chạy ngầm hoặc chờ quét QR không
async function checkActiveZaloDownload() {
    try {
        const res = await fetch('/api/zalo/download/progress/poll');
        const data = await res.json();
        if (data.success && data.data) {
            const p = data.data;
            if (p.status && p.status !== 'idle' && p.status !== 'done' && p.status !== 'error') {
                console.log('[Zalo] Phát hiện tiến trình tải đang chạy ngầm, tự động kết nối:', p.status);
                zaloIsDownloading = true;
                const cancelBtn = document.getElementById('btn-cancel-zalo-download');
                if (cancelBtn) cancelBtn.style.display = 'inline-flex';
                listenZaloProgress();
            }
        }
    } catch (e) {
        console.warn('checkActiveZaloDownload error:', e);
    }
}

// Hủy / Dừng tiến trình tải ảnh đang chạy
async function cancelZaloDownload() {
    if (!await confirmAction('Bạn có chắc chắn muốn hủy / dừng tiến trình tải ảnh hiện tại không?')) return;
    try {
        appendZaloLog('[Hệ thống] Đang gửi yêu cầu dừng tiến trình...', 'warning');
        const res = await fetch('/api/zalo/download/cancel', { method: 'POST' });
        const data = await res.json();
        showToast(data.message || 'Đã hủy tiến trình tải ảnh', 'info');
        appendZaloLog('[Hệ thống] ' + (data.message || 'Đã hủy tiến trình tải ảnh.'), 'info');
    } catch (e) {
        showToast('Lỗi khi hủy tiến trình: ' + e.message, 'error');
    } finally {
        zaloIsDownloading = false;
        zaloCurrentQrSource = null;
        if (zaloEventSource) {
            zaloEventSource.close();
            zaloEventSource = null;
        }
        const cancelBtn = document.getElementById('btn-cancel-zalo-download');
        if (cancelBtn) cancelBtn.style.display = 'none';

        const btn = document.getElementById('btn-start-zalo-download');
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
                <span>Bắt Đầu Tải Ảnh Về Máy</span>`;
        }

        const qrCard = document.getElementById('zalo-qr-card');
        if (qrCard) qrCard.style.display = 'none';
        const inlineBox = document.getElementById('zalo-download-qr-box');
        if (inlineBox) inlineBox.style.display = 'none';

        updateZaloProgressUI({
            status: 'idle',
            total: 0,
            downloaded: 0,
            skipped: 0,
            failed: 0,
            currentFile: 'Đã dừng tiến trình.'
        });
    }
}

function restoreZaloGroupSelection() {
    // Restore selected group card UI if a group was previously selected
    if (selectedZaloGroupId && selectedZaloGroupName) {
        const card = document.getElementById('zalo-selected-group-card');
        const trigger = document.getElementById('zalo-group-trigger');
        const nameEl = document.getElementById('zalo-selected-name');
        const memEl = document.getElementById('zalo-selected-members');
        const subEl = document.getElementById('zalo-selected-sub');

        if (nameEl) nameEl.textContent = selectedZaloGroupName;
        if (memEl) memEl.textContent = `${selectedZaloGroupMembers} thành viên`;
        if (subEl) subEl.textContent = `Mã nhóm ID: ${selectedZaloGroupId}`;
        if (card) card.style.display = 'flex';
        if (trigger) trigger.style.display = 'none';
    }
}

// Check Zalo status & update UI
async function checkZaloStatus(showToastMsg = false) {
    const badge = document.getElementById('zalo-conn-status-badge');
    const sideBadge = document.getElementById('zalo-sidebar-badge');
    const userPill = document.getElementById('zalo-user-pill');
    const btnLogin = document.getElementById('btn-zalo-login-qr');
    const btnLogout = document.getElementById('btn-zalo-logout');
    const subtitle = document.getElementById('zalo-conn-subtitle');
    const qrCard = document.getElementById('zalo-qr-card');

    try {
        const res = await fetch('/api/zalo/status');
        const data = await res.json();

        if (data.success && data.data) {
            const status = data.data;
            if (status.isLoggedIn) {
                if (badge) {
                    badge.className = 'zalo-badge online';
                    badge.textContent = 'Đã kết nối';
                }
                if (sideBadge) {
                    sideBadge.className = 'zalo-nav-badge online';
                }
                if (userPill) {
                    userPill.style.display = 'flex';
                    const nameEl = document.getElementById('zalo-user-name');
                    const uidEl = document.getElementById('zalo-user-uid');
                    const avtEl = document.getElementById('zalo-user-avatar');
                    if (status.qrUserInfo) {
                        if (nameEl) nameEl.textContent = status.qrUserInfo.display_name || 'Người dùng Zalo';
                        if (uidEl) uidEl.textContent = 'UID: ' + (status.qrUserInfo.uid || 'Zalo Web');
                        if (avtEl && status.qrUserInfo.avatar) avtEl.src = status.qrUserInfo.avatar;
                    } else {
                        if (nameEl) nameEl.textContent = 'Tài khoản Zalo';
                        if (uidEl) uidEl.textContent = 'Đã kết nối phiên làm việc';
                        if (avtEl) avtEl.src = 'https://chat.zalo.me/favicon.ico';
                    }
                }
                if (btnLogin) btnLogin.style.display = 'none';
                if (btnLogout) btnLogout.style.display = 'inline-flex';
                if (qrCard) qrCard.style.display = 'none';
                if (subtitle) subtitle.textContent = 'Tài khoản đã sẵn sàng tải ảnh từ các nhóm chấm công.';

                // Automatically load groups if not loaded yet
                if (zaloGroupsList.length === 0) {
                    loadZaloGroups();
                }

                if (showToastMsg) showToast('Zalo đã kết nối thành công!', 'success');
            } else {
                if (badge) {
                    if (status.loginInProgress || status.qrStatus === 'waiting_scan' || status.qrStatus === 'scanned') {
                        badge.className = 'zalo-badge waiting';
                        badge.textContent = 'Đang chờ quét QR...';
                        if (sideBadge) sideBadge.className = 'zalo-nav-badge warning';
                    } else {
                        badge.className = 'zalo-badge offline';
                        badge.textContent = 'Chưa kết nối';
                        if (sideBadge) sideBadge.className = 'zalo-nav-badge';
                    }
                }
                if (userPill) userPill.style.display = 'none';
                if (btnLogin) btnLogin.style.display = 'inline-flex';
                if (btnLogout) btnLogout.style.display = 'none';
                if (subtitle) subtitle.textContent = 'Chưa đăng nhập. Vui lòng bấm "Đăng Nhập Zalo (QR)" để kết nối tài khoản Zalo.';
                if (showToastMsg) showToast('Chưa kết nối Zalo. Vui lòng đăng nhập.', 'warning');
            }
        } else {
            throw new Error(data.error || 'Không thể kết nối Zalo Service');
        }
    } catch (err) {
        if (badge) {
            badge.className = 'zalo-badge offline';
            badge.textContent = 'Dịch vụ Zalo Chưa Bật';
        }
        if (sideBadge) sideBadge.className = 'zalo-nav-badge';
        if (subtitle) subtitle.textContent = 'Dịch vụ Zalo Service chưa phản hồi. Nhấn "Khởi Động Lại Service" để bật lại.';
        if (showToastMsg) showToast('Lỗi kết nối Zalo Service: ' + err.message, 'error');
    }
}

// Khởi động lại riêng Zalo Service (Port 3001)
async function restartZaloServiceUi() {
    const btn = document.getElementById('btn-zalo-restart');
    const badge = document.getElementById('zalo-conn-status-badge');
    const subtitle = document.getElementById('zalo-conn-subtitle');
    const qrCard = document.getElementById('zalo-qr-card');

    if (btn && btn.classList.contains('is-loading')) return;

    if (btn) {
        btn.classList.add('is-loading');
        btn.disabled = true;
    }
    if (badge) {
        badge.className = 'zalo-badge waiting';
        badge.textContent = 'Đang khởi động lại...';
    }
    if (subtitle) {
        subtitle.textContent = 'Đang dọn dẹp port 3001 và khởi động lại Zalo Service (Node.js)...';
    }
    if (qrCard) {
        qrCard.style.display = 'none';
    }
    appendZaloLog('[Hệ thống] Đang yêu cầu khởi động lại Zalo Service...', 'info');
    showToast('Đang khởi động lại Zalo Service (Port 3001)...', 'info');

    try {
        const res = await fetch('/api/zalo/restart', { method: 'POST' });
        const data = await res.json();

        if (res.ok && data.success) {
            showToast('✓ ' + (data.message || 'Zalo Service đã khởi động lại thành công!'), 'success');
            appendZaloLog('[Hệ thống] ✓ Zalo Service đã khởi động lại và sẵn sàng!', 'success');
            await checkZaloStatus(false);
        } else {
            throw new Error(data.error || 'Khởi động lại Zalo Service thất bại');
        }
    } catch (err) {
        showToast('✗ Lỗi khởi động lại: ' + err.message, 'error');
        appendZaloLog('[Lỗi] Không thể khởi động lại Zalo Service: ' + err.message, 'error');
        if (badge) {
            badge.className = 'zalo-badge offline';
            badge.textContent = 'Khởi Động Thất Bại';
        }
        if (subtitle) {
            subtitle.textContent = 'Không thể khởi động lại Zalo Service. Vui lòng kiểm tra Node.js hoặc xem zalo_service.log.';
        }
    } finally {
        if (btn) {
            btn.classList.remove('is-loading');
            btn.disabled = false;
        }
    }
}

// Bắt đầu đăng nhập bằng QR
async function startZaloQrLogin(force = false) {
    zaloCurrentQrSource = 'login';
    const qrCard = document.getElementById('zalo-qr-card');
    const qrLoading = document.getElementById('zalo-qr-loading');
    const qrImg = document.getElementById('zalo-qr-img');
    const qrMsg = document.getElementById('zalo-qr-status-msg');

    if (qrCard) qrCard.style.display = 'block';
    if (qrLoading) qrLoading.style.display = 'flex';
    if (qrImg) {
        qrImg.style.display = 'none';
        qrImg.src = '';
    }
    if (qrMsg) qrMsg.innerHTML = '<span class="pulse-dot warning"></span> ' + (force ? 'Đang làm mới và tạo mã QR mới...' : 'Đang kết nối Zalo & tạo mã QR...');

    appendZaloLog('[Hệ thống] ' + (force ? 'Đang làm mới và tạo mã QR đăng nhập mới...' : 'Bắt đầu tạo mã QR đăng nhập Zalo...'), 'info');

    try {
        const res = await fetch('/api/zalo/login/qr', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ force: !!force })
        });
        const data = await res.json();
        if (!data.success && data.error) {
            showToast('Lỗi tạo QR: ' + data.error, 'error');
            appendZaloLog('[Lỗi] ' + data.error, 'error');
            return;
        }

        // Bắt đầu polling QR status và QR image
        if (zaloPollTimer) clearInterval(zaloPollTimer);
        pollZaloQr();
        zaloPollTimer = setInterval(pollZaloQr, 2000);

    } catch (err) {
        showToast('Lỗi gửi yêu cầu tạo QR: ' + err.message, 'error');
        appendZaloLog('[Lỗi] ' + err.message, 'error');
    }
}

// Hiển thị trực tiếp cửa sổ Chrome thật
async function showZaloBrowserWindow() {
    try {
        appendZaloLog('[Hệ thống] Đang yêu cầu hiển thị cửa sổ trình duyệt Chrome...', 'info');
        const res = await fetch('/api/zalo/download/browser/show', { method: 'POST' });
        const data = await res.json();
        if (data.success) {
            showToast('Đã mở cửa sổ Chrome trên màn hình', 'success');
            appendZaloLog('[Hệ thống] Đã đưa cửa sổ Chrome lên màn hình.', 'success');
        } else {
            showToast('Không thể mở cửa sổ Chrome: ' + (data.error || 'Trình duyệt chưa sẵn sàng'), 'warning');
            appendZaloLog('[Cảnh báo] ' + (data.error || 'Trình duyệt chưa sẵn sàng'), 'warning');
        }
    } catch (err) {
        showToast('Lỗi kết nối: ' + err.message, 'error');
        appendZaloLog('[Lỗi] ' + err.message, 'error');
    }
}

// Xử lý làm mới mã QR thông minh theo ngữ cảnh (Tải ảnh vs Đăng nhập API)
async function handleZaloQrRefresh() {
    if (zaloIsDownloading || zaloCurrentQrSource === 'download') {
        const qrLoading = document.getElementById('zalo-qr-loading');
        const qrImg = document.getElementById('zalo-qr-img');
        const qrMsg = document.getElementById('zalo-qr-status-msg');

        if (qrLoading) qrLoading.style.display = 'flex';
        if (qrImg) qrImg.style.display = 'none';
        if (qrMsg) qrMsg.innerHTML = '<span class="pulse-dot warning"></span> Đang tạo lại mã QR tải ảnh mới...';

        appendZaloLog('[Hệ thống] Đang yêu cầu trình duyệt làm mới mã QR tải ảnh...', 'info');
        try {
            const res = await fetch('/api/zalo/download/qr/refresh', { method: 'POST' });
            const data = await res.json();
            if (data.success && data.data && data.data.image) {
                if (qrImg) {
                    qrImg.src = data.data.image;
                    qrImg.style.display = 'block';
                }
                if (qrLoading) qrLoading.style.display = 'none';
                if (qrMsg) qrMsg.innerHTML = '<span class="status-indicator-dot warning"></span> <strong>Đã cập nhật mã QR mới!</strong> Vui lòng quét bằng app Zalo.';
                appendZaloLog('[Hệ thống] Đã tạo mã QR mới thành công.', 'success');
            } else {
                throw new Error(data.error || 'Không thể tạo mã mới');
            }
        } catch (err) {
            showToast('Lỗi làm mới QR: ' + err.message, 'error');
            appendZaloLog('[Lỗi] ' + err.message, 'error');
            if (qrLoading) qrLoading.style.display = 'none';
            if (qrImg) qrImg.style.display = 'block';
        }
    } else {
        await startZaloQrLogin(true);
    }
}

// Poll QR image and status
async function pollZaloQr() {
    const qrLoading = document.getElementById('zalo-qr-loading');
    const qrImg = document.getElementById('zalo-qr-img');
    const qrMsg = document.getElementById('zalo-qr-status-msg');

    try {
        // Lấy QR image
        const imgRes = await fetch('/api/zalo/login/qr/image');
        const imgData = await imgRes.json();
        if (imgData.success && imgData.data && imgData.data.image) {
            let src = imgData.data.image;
            if (!src.startsWith('data:image/') && !src.startsWith('http')) {
                src = 'data:image/png;base64,' + src;
            }
            if (qrImg.src !== src) {
                qrImg.src = src;
            }
            qrImg.onerror = () => {
                console.warn('QR base64 failed, falling back to /api/zalo/login/qr/raw...');
                qrImg.onerror = null;
                qrImg.src = `/api/zalo/login/qr/raw?t=${Date.now()}`;
            };
            qrImg.style.display = 'block';
            if (qrLoading) qrLoading.style.display = 'none';
        }

        // Lấy trạng thái đăng nhập
        const stRes = await fetch('/api/zalo/login/qr/status');
        const stData = await stRes.json();
        if (stData.success && stData.data) {
            const st = stData.data;
            if (st.status === 'scanned') {
                if (qrMsg) qrMsg.innerHTML = '<span class="status-indicator-dot success"></span> <strong>Đã quét mã!</strong> Vui lòng bấm Xác Nhận Đăng Nhập trên điện thoại...';
                appendZaloLog('[Zalo] Điện thoại đã quét mã QR, chờ xác nhận...', 'info');
            } else if (st.isLoggedIn) {
                // Đăng nhập thành công!
                if (zaloPollTimer) {
                    clearInterval(zaloPollTimer);
                    zaloPollTimer = null;
                }
                const qrCard = document.getElementById('zalo-qr-card');
                if (qrCard) qrCard.style.display = 'none';
                showToast('Đăng nhập Zalo thành công!', 'success');
                appendZaloLog('[Zalo] Đăng nhập thành công! Đang tải danh sách nhóm...', 'success');
                await checkZaloStatus();
                await loadZaloGroups();
            } else if (st.status === 'declined') {
                if (qrMsg) qrMsg.innerHTML = '<span class="status-indicator-dot danger"></span> Bạn đã từ chối đăng nhập trên điện thoại.';
                appendZaloLog('[Zalo] Đã từ chối đăng nhập trên điện thoại.', 'warning');
                if (zaloPollTimer) {
                    clearInterval(zaloPollTimer);
                    zaloPollTimer = null;
                }
            } else if (st.status === 'error') {
                if (qrMsg) qrMsg.textContent = st.error || 'Đăng nhập Zalo Web không thành công';
                if (qrImg) qrImg.style.display = 'none';
                if (qrLoading) qrLoading.style.display = 'none';
                if (zaloPollTimer) clearInterval(zaloPollTimer);
                zaloPollTimer = null;
            } else if (st.status === 'expired') {
                if (qrMsg) qrMsg.innerHTML = '<span class="status-indicator-dot warning"></span> Mã QR đã hết hạn, đang tạo lại...';
                if (qrImg) qrImg.style.display = 'none';
                if (qrLoading) qrLoading.style.display = 'flex';
            }
        }
    } catch (err) {
        console.error('Lỗi poll QR:', err);
    }
}

async function cancelZaloLogin() {
    if (zaloPollTimer) {
        clearInterval(zaloPollTimer);
        zaloPollTimer = null;
    }
    const qrCard = document.getElementById('zalo-qr-card');
    if (qrCard) qrCard.style.display = 'none';

    if (zaloIsDownloading || zaloCurrentQrSource === 'download') {
        try {
            await fetch('/api/zalo/download/cancel', { method: 'POST' });
        } catch {}
        zaloIsDownloading = false;
        zaloCurrentQrSource = null;
        if (zaloEventSource) {
            zaloEventSource.close();
            zaloEventSource = null;
        }
        const btn = document.getElementById('btn-start-zalo-download');
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
                <span>Bắt Đầu Tải Ảnh Về Máy</span>`;
        }
    }
    appendZaloLog('[Hệ thống] Đã đóng cửa sổ đăng nhập / hủy chờ quét QR.', 'info');
}

async function logoutZalo() {
    if (!await confirmAction('Bạn có chắc chắn muốn đăng xuất tài khoản Zalo không?')) return;
    try {
        const res = await fetch('/api/zalo/logout', { method: 'POST' });
        const data = await res.json();
        if (data.success) {
            showToast('Đã đăng xuất Zalo', 'info');
            appendZaloLog('[Zalo] Đã đăng xuất thành công.', 'info');
            zaloGroupsList = [];
            const sel = document.getElementById('zalo-group-select');
            if (sel) sel.innerHTML = '<option value="" disabled selected>-- Cần đăng nhập Zalo để tải danh sách nhóm --</option>';
            checkZaloStatus();
        } else {
            showToast(data.error || 'Lỗi khi đăng xuất', 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

let selectedZaloGroupId = '';
let selectedZaloGroupName = '';
let selectedZaloGroupMembers = 0;
let lastZaloDownloadedProject = 'Dự án Long An';

function toggleZaloGroupDropdown(forceOpen) {
    const menu = document.getElementById('zalo-group-dropdown-menu');
    const searchInput = document.getElementById('zalo-group-search');
    if (!menu) return;

    const isOpen = menu.style.display === 'block';
    const shouldOpen = forceOpen !== undefined ? forceOpen : !isOpen;

    if (shouldOpen) {
        menu.style.display = 'block';
        if (searchInput) {
            searchInput.value = '';
            filterZaloGroups();
            setTimeout(() => searchInput.focus(), 50);
        }
    } else {
        menu.style.display = 'none';
    }
}

function clearZaloGroupSearch() {
    const searchInput = document.getElementById('zalo-group-search');
    if (searchInput) {
        searchInput.value = '';
        filterZaloGroups();
        searchInput.focus();
    }
}

function escapeHtml(str) {
    if (!str) return '';
    return String(str)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
}

function renderZaloGroupItems(groups) {
    const listEl = document.getElementById('zalo-group-items-list');
    const countEl = document.getElementById('zalo-group-count-text');
    if (!listEl) return;
    listEl.innerHTML = '';

    if (countEl) {
        countEl.textContent = `${groups.length} / ${zaloGroupsList.length} nhóm`;
    }

    if (groups.length === 0) {
        listEl.innerHTML = '<div class="zalo-dropdown-empty">Không tìm thấy nhóm Zalo nào phù hợp</div>';
        return;
    }

    groups.forEach(g => {
        const item = document.createElement('div');
        const isSelected = g.id === selectedZaloGroupId;
        item.className = `zalo-group-item ${isSelected ? 'selected' : ''}`;
        
        const firstLetter = (g.name || 'Z').trim().charAt(0).toUpperCase();
        const avatarHtml = g.avatar
            ? `<img src="${escapeHtml(g.avatar)}" class="zalo-item-avatar" alt="" onerror="this.style.display='none';this.nextElementSibling.style.display='flex'">`
              + `<span class="zalo-item-icon" style="display:none">${escapeHtml(firstLetter)}</span>`
            : `<span class="zalo-item-icon">${escapeHtml(firstLetter)}</span>`;
        const memberText = g.totalMember ? `${g.totalMember} thành viên` : '';

        item.innerHTML = `
            <div class="zalo-item-name">
                ${avatarHtml}
                <span>${escapeHtml(g.name)}</span>
            </div>
            <span class="zalo-item-count">${memberText}</span>
        `;

        item.onclick = () => {
            selectZaloGroup(g.id, g.name, g.totalMember || 0);
        };

        listEl.appendChild(item);
    });
}

function selectZaloGroup(groupId, groupName, memberCount) {
    selectedZaloGroupId = groupId;
    selectedZaloGroupName = groupName;
    selectedZaloGroupMembers = memberCount;

    // Cập nhật input ẩn
    const hidId = document.getElementById('zalo-selected-group-id');
    const hidTitle = document.getElementById('zalo-selected-group-title');
    if (hidId) hidId.value = groupId;
    if (hidTitle) hidTitle.value = groupName;

    // Cập nhật giao diện xác nhận (Confirmation Card)
    const card = document.getElementById('zalo-selected-group-card');
    const trigger = document.getElementById('zalo-group-trigger');
    const nameEl = document.getElementById('zalo-selected-name');
    const memEl = document.getElementById('zalo-selected-members');
    const subEl = document.getElementById('zalo-selected-sub');

    if (nameEl) nameEl.textContent = groupName;
    if (memEl) memEl.textContent = `${memberCount} thành viên`;
    if (subEl) subEl.textContent = `Mã nhóm ID: ${groupId}`;

    // ĐÓNG DROPDOWN VÀ HIỂN THỊ XÁC NHẬN
    if (card) card.style.display = 'flex';
    if (trigger) trigger.style.display = 'none';
    toggleZaloGroupDropdown(false);

    // Tự động tìm dự án khớp tên
    const projectSel = document.getElementById('zalo-target-project');
    if (projectSel) {
        for (let i = 0; i < projectSel.options.length; i++) {
            if (projectSel.options[i].text.toLowerCase().includes(groupName.toLowerCase()) ||
                projectSel.options[i].value.toLowerCase() === groupName.toLowerCase()) {
                projectSel.selectedIndex = i;
                break;
            }
        }
    }

    appendZaloLog(`[Đã chọn nhóm] "${groupName}" (${memberCount} thành viên)`, 'success');
    showToast(`Đã chọn nhóm: ${groupName}`, 'success');

    renderZaloGroupItems(zaloGroupsList);
}

function filterZaloGroups() {
    const query = (document.getElementById('zalo-group-search')?.value || '').toLowerCase().trim();
    if (!query) {
        renderZaloGroupItems(zaloGroupsList);
        return;
    }
    const filtered = zaloGroupsList.filter(g => (g.name || '').toLowerCase().includes(query));
    renderZaloGroupItems(filtered);
}

// Load groups from Zalo
async function loadZaloGroups(manual = false) {
    const listEl = document.getElementById('zalo-group-items-list');
    const triggerText = document.getElementById('zalo-group-trigger-text');

    if (manual) appendZaloLog('[Zalo] Đang lấy danh sách nhóm từ Zalo...', 'info');
    if (listEl) listEl.innerHTML = '<div class="zalo-dropdown-empty">Đang tải danh sách nhóm...</div>';

    try {
        const res = await fetch('/api/zalo/groups');
        const data = await res.json();

        if (data.success && Array.isArray(data.data)) {
            zaloGroupsList = data.data;
            renderZaloGroupItems(zaloGroupsList);
            if (triggerText && !selectedZaloGroupId) {
                triggerText.textContent = `-- Chọn trong ${zaloGroupsList.length} nhóm Zalo --`;
            }
            if (manual) showToast(`Đã tải ${zaloGroupsList.length} nhóm Zalo`, 'success');
            appendZaloLog(`[Zalo] Đã nạp thành công ${zaloGroupsList.length} nhóm.`, 'success');
        } else {
            if (listEl) listEl.innerHTML = `<div class="zalo-dropdown-empty">${data.error || 'Chưa có nhóm nào hoặc chưa đăng nhập'}</div>`;
            if (manual) showToast(data.error || 'Chưa thể tải nhóm (hãy kiểm tra đăng nhập)', 'warning');
        }
    } catch (err) {
        if (listEl) listEl.innerHTML = `<div class="zalo-dropdown-empty">Lỗi kết nối: ${err.message}</div>`;
    }
}

// Load projects into Zalo target dropdown
async function loadProjectsForZalo() {
    const projectSel = document.getElementById('zalo-target-project');
    if (!projectSel) return;

    try {
        const res = await fetch('/api/projects');
        const data = await res.json();
        projectSel.innerHTML = '<option value="__AUTO__">[Tự động tạo theo tên nhóm Zalo]</option>';

        if (data.success && Array.isArray(data.projects)) {
            data.projects.forEach(p => {
                const opt = document.createElement('option');
                opt.value = p.name;
                opt.textContent = `Dự án: ${p.name}`;
                projectSel.appendChild(opt);
            });
        }
    } catch (err) {
        console.error('Lỗi nạp dự án cho Zalo:', err);
    }
}

// Set Date Presets
function setZaloDatePreset(preset) {
    const fromInput = document.getElementById('zalo-date-from');
    const toInput = document.getElementById('zalo-date-to');
    if (!fromInput || !toInput) return;

    const formatDate = (d) => {
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    };

    const today = new Date();
    const todayStr = formatDate(today);

    document.querySelectorAll('.pill-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.preset === preset);
    });

    if (preset === 'today') {
        setDatePickerValue(fromInput, todayStr);
        setDatePickerValue(toInput, todayStr);
    } else if (preset === 'yesterday') {
        const y = new Date();
        y.setDate(y.getDate() - 1);
        const yStr = formatDate(y);
        setDatePickerValue(fromInput, yStr);
        setDatePickerValue(toInput, yStr);
    } else if (preset === '3days') {
        const d = new Date();
        d.setDate(d.getDate() - 2);
        setDatePickerValue(fromInput, formatDate(d));
        setDatePickerValue(toInput, todayStr);
    } else if (preset === '7days') {
        const d = new Date();
        d.setDate(d.getDate() - 6);
        setDatePickerValue(fromInput, formatDate(d));
        setDatePickerValue(toInput, todayStr);
    } else if (preset === 'all') {
        setDatePickerValue(fromInput, '');
        setDatePickerValue(toInput, '');
    }
}

// Start downloading images from group
async function startZaloDownload() {
    if (zaloIsDownloading) {
        showToast('Đang có tiến trình tải ảnh đang chạy!', 'warning');
        return;
    }

    const groupId = document.getElementById('zalo-selected-group-id')?.value || selectedZaloGroupId;
    const groupNameOriginal = document.getElementById('zalo-selected-group-title')?.value || selectedZaloGroupName;

    if (!groupId) {
        showToast('Vui lòng chọn 1 nhóm Zalo để tải ảnh!', 'warning');
        toggleZaloGroupDropdown(true);
        return;
    }

    const targetProjectSel = document.getElementById('zalo-target-project');
    let targetProjectName = groupNameOriginal;
    if (targetProjectSel && targetProjectSel.value && targetProjectSel.value !== '__AUTO__') {
        targetProjectName = targetProjectSel.value;
    }
    lastZaloDownloadedProject = targetProjectName;

    // Tự động đảm bảo dự án được tạo đầy đủ trên hệ thống (bao gồm thư mục Ảnh BV và input_images 31 ngày)
    try {
        await apiPost('/api/projects/create', { name: targetProjectName });
        if (typeof loadProjects === 'function') {
            await loadProjects();
        }
    } catch (e) {
        console.warn('Tạo cấu trúc dự án tự động:', e);
    }

    const dateFrom = document.getElementById('zalo-date-from')?.value || null;
    const dateTo = document.getElementById('zalo-date-to')?.value || null;
    const folderFormat = document.getElementById('zalo-folder-format')?.value || 'date';
    const msgLimit = parseInt(document.getElementById('zalo-msg-limit')?.value || '0', 10);

    const btn = document.getElementById('btn-start-zalo-download');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<div class="spinner spinner-sm"></div> <span>Đang Quét & Tải Ảnh...</span>';
    }
    zaloIsDownloading = true;

    // Reset progress UI
    updateZaloProgressUI({
        total: 0,
        downloaded: 0,
        skipped: 0,
        failed: 0,
        status: 'downloading',
        currentFile: 'Đang gửi yêu cầu tới Zalo...'
    });
    const banner = document.getElementById('zalo-completion-banner');
    if (banner) banner.style.display = 'none';

    appendZaloLog(`[Tải ảnh] Bắt đầu tải cho nhóm "${groupNameOriginal}" -> Thư mục dự án "${targetProjectName}"`, 'info');
    if (dateFrom || dateTo) {
        appendZaloLog(`[Lọc ngày] Từ: ${dateFrom || 'Đầu'} đến: ${dateTo || 'Nay'}`, 'info');
    }

    try {
        const showBrowser = document.getElementById('zalo-show-browser')?.checked || false;
        const payload = {
            groupName: groupNameOriginal,
            projectName: targetProjectName,
            dateFrom: dateFrom,
            dateTo: dateTo,
            folderFormat: 'YYYY-MM-DD',
            count: msgLimit,
            headless: !showBrowser
        };

        const res = await fetch(`/api/zalo/groups/${groupId}/download`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const data = await res.json();
        if (!data.success) {
            if (data.error && data.error.includes('tiến trình tải ảnh đang chạy')) {
                showToast('Đang kết nối vào tiến trình tải đang chạy...', 'info');
                appendZaloLog('[Hệ thống] Đang có tiến trình tải ảnh đang hoạt động. Đã tự động kết nối theo dõi...', 'info');
                zaloIsDownloading = true;
                const cancelBtn = document.getElementById('btn-cancel-zalo-download');
                if (cancelBtn) cancelBtn.style.display = 'inline-flex';
                listenZaloProgress();
                return;
            }
            throw new Error(data.error || 'Lỗi bắt đầu tải');
        }

        const cancelBtn = document.getElementById('btn-cancel-zalo-download');
        if (cancelBtn) cancelBtn.style.display = 'inline-flex';
        appendZaloLog(`[Zalo] ${data.message || 'Đã tìm thấy tin nhắn, đang tải ảnh...'}`, 'info');

        // Lắng nghe tiến trình tải
        listenZaloProgress();

    } catch (err) {
        showToast('Lỗi tải ảnh: ' + err.message, 'error');
        appendZaloLog('[Lỗi] ' + err.message, 'error');
        zaloIsDownloading = false;
        const cancelBtn = document.getElementById('btn-cancel-zalo-download');
        if (cancelBtn) cancelBtn.style.display = 'none';
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `
                <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7 10 12 15 17 10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
                <span>Bắt Đầu Tải Ảnh Về Máy</span>`;
        }
    }
}

// Listen to download progress using SSE & fallback polling
function listenZaloProgress() {
    if (zaloEventSource) {
        zaloEventSource.close();
        zaloEventSource = null;
    }

    let isDone = false;
    let fallbackPollTimer = null;

    let lastLogIndex = 0;
    const onProgressData = (p) => {
        updateZaloProgressUI(p);

        // Đồng bộ nhật ký thời gian thực từ tiến trình tải
        if (p.log && Array.isArray(p.log)) {
            const logOffset = p.logOffset || 0;
            lastLogIndex = Math.max(lastLogIndex, logOffset);
            while (lastLogIndex < logOffset + p.log.length) {
                const logEntry = p.log[lastLogIndex++ - logOffset];
                let logType = 'info';
                if (logEntry.includes('[Lỗi]') || logEntry.includes('[✗]') || logEntry.includes('Error')) {
                    logType = 'error';
                } else if (logEntry.includes('[Cảnh báo]')) {
                    logType = 'warning';
                } else if (logEntry.includes('[✓]') || logEntry.includes('Hoàn thành') || logEntry.includes('thành công')) {
                    logType = 'success';
                }
                const term = document.getElementById('zalo-log-terminal');
                if (term) {
                    const line = document.createElement('div');
                    line.className = `log-line ${logType}`;
                    line.textContent = logEntry;
                    term.appendChild(line);
                    term.scrollTop = term.scrollHeight;
                }
            }
        } else if (p.currentFile) {
            appendZaloLog(`[Đang lưu] ${p.currentFile}`, 'info');
        }

        // Tự động hiển thị mã QR ngay trên giao diện web khi trình duyệt yêu cầu quét mã
        const qrCard = document.getElementById('zalo-qr-card');
        const qrImg = document.getElementById('zalo-qr-img');
        const qrLoading = document.getElementById('zalo-qr-loading');
        const qrMsg = document.getElementById('zalo-qr-status-msg');

        if (p.status === 'waiting_qr' || p.qrImage) {
            zaloCurrentQrSource = 'download';
            if (qrCard && qrCard.style.display !== 'block') {
                qrCard.style.display = 'block';
                qrCard.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
            if (p.qrImage && qrImg) {
                if (qrImg.src !== p.qrImage) {
                    qrImg.src = p.qrImage;
                }
                qrImg.style.display = 'block';
                if (qrLoading) qrLoading.style.display = 'none';
            }
            if (qrMsg) {
                let countdownHtml = '';
                if (p.qrExpiresAt) {
                    const remaining = Math.max(0, Math.ceil((p.qrExpiresAt - Date.now()) / 1000));
                    if (remaining > 0) {
                        const color = remaining <= 15 ? '#ff4444' : remaining <= 30 ? '#ffaa00' : '#4caf50';
                        countdownHtml = ` <span style="color:${color};font-weight:bold;font-size:0.95em;">(⏱ còn ${remaining}s)</span>`;
                    } else {
                        countdownHtml = ' <span style="color:#ff4444;font-weight:bold;">⏳ Đang lấy mã mới...</span>';
                    }
                }
                qrMsg.innerHTML = '<span class="status-indicator-dot warning"></span> <strong>Quét mã QR dưới đây trong app Zalo để bắt đầu tải ảnh</strong>' + countdownHtml + ' (chỉ cần làm 1 lần, tự động làm mới khi hết hạn)...';
            }
        } else if (p.status && p.status !== 'waiting_qr' && p.status !== 'idle' && !p.qrImage) {
            if (qrCard && qrCard.style.display !== 'none') {
                qrCard.style.display = 'none';
            }
            if (zaloCurrentQrSource === 'download') {
                zaloCurrentQrSource = null;
                setTimeout(() => {
                    checkZaloStatus(false);
                    loadZaloGroups(false);
                }, 1500);
            }
        }

        // Đồng thời hiển thị khung QR trực quan ngay tại cột Tiến Trình Tải
        const inlineBox = document.getElementById('zalo-download-qr-box');
        const inlineImg = document.getElementById('zalo-inline-qr-img');
        const inlineCountdown = document.getElementById('zalo-inline-qr-countdown');

        if (p.status === 'waiting_qr' || p.qrImage) {
            if (inlineBox) inlineBox.style.display = 'block';
            if (inlineImg && p.qrImage) {
                if (inlineImg.src !== p.qrImage) inlineImg.src = p.qrImage;
            }
            if (inlineCountdown && p.qrExpiresAt) {
                const remaining = Math.max(0, Math.ceil((p.qrExpiresAt - Date.now()) / 1000));
                inlineCountdown.innerHTML = remaining > 0
                    ? `⏱ Mã QR còn hiệu lực: <strong style="color:#16a34a;">${remaining}s</strong>`
                    : `⏳ Đang tự động đổi mã mới...`;
            }
        } else {
            if (inlineBox) inlineBox.style.display = 'none';
        }

        const cancelBtn = document.getElementById('btn-cancel-zalo-download');
        if (cancelBtn) {
            if (['downloading', 'waiting_qr', 'opening_group', 'scanning_media'].includes(p.status)) {
                cancelBtn.style.display = 'inline-flex';
            } else if (p.status === 'done' || p.status === 'error' || p.status === 'idle') {
                cancelBtn.style.display = 'none';
            }
        }

        const startBtn = document.getElementById('btn-start-zalo-download');
        if (startBtn && p.status === 'waiting_qr') {
            startBtn.disabled = true;
            startBtn.innerHTML = '<div class="spinner spinner-sm"></div> <span>Đang Chờ Quét QR...</span>';
        }

        if (p.errors && p.errors.length > 0) {
            const lastErr = p.errors[p.errors.length - 1];
            appendZaloLog(`[Lỗi file] ${lastErr}`, 'error');
        }

        if (p.status === 'done' || p.status === 'error') {
            if (!isDone) {
                isDone = true;
                zaloIsDownloading = false;
                if (cancelBtn) cancelBtn.style.display = 'none';
                if (zaloEventSource) {
                    zaloEventSource.close();
                    zaloEventSource = null;
                }
                if (fallbackPollTimer) {
                    clearInterval(fallbackPollTimer);
                    fallbackPollTimer = null;
                }

                const btn = document.getElementById('btn-start-zalo-download');
                if (btn) {
                    btn.disabled = false;
                    btn.innerHTML = `
                        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                            <polyline points="7 10 12 15 17 10"></polyline>
                            <line x1="12" y1="15" x2="12" y2="3"></line>
                        </svg>
                        <span>Bắt Đầu Tải Ảnh Về Máy</span>`;
                }

                if (p.status === 'done') {
                    const actualDownloaded = p.counterVersion === 2 ? (p.downloaded || 0) : Math.max(0, (p.downloaded || 0) - (p.skipped || 0));
                    const scanWarning = p.scanStopReason === 'limit' ? ' Chưa quét hết kho ảnh vì đã đạt giới hạn đã chọn.' : '';
                    const historyWarning = p.historyNotice ? ` Thông báo Zalo: ${p.historyNotice}` : '';
                    const warningText = `Ảnh xem trước: ${p.lowQuality || 0}; không rõ ngày gửi (đã lưu vào thư mục unknown-date): ${p.unknownDate || 0}; lỗi tải: ${p.failed || 0}.${scanWarning}${historyWarning}`;
                    const hasWarnings = (p.lowQuality || 0) + (p.unknownDate || 0) + (p.failed || 0) > 0 || !!scanWarning || !!historyWarning;
                    appendZaloLog(`[Hoàn thành] Đã lưu ${actualDownloaded} ảnh mới. ${warningText}`, hasWarnings ? 'warning' : 'success');
                    showToast(`Đã xử lý xong: ${actualDownloaded} ảnh mới.${hasWarnings ? ' Có cảnh báo, xem kết quả tải.' : ''}`, hasWarnings ? 'warning' : 'success');
                    
                    // Nạp lại danh sách dự án và tự động chọn dự án vừa tải
                    if (typeof loadProjects === 'function') {
                        loadProjects().then(() => {
                            if (lastZaloDownloadedProject && typeof onProjectChange === 'function') {
                                onProjectChange(lastZaloDownloadedProject);
                            }
                        });
                    }

                    const banner = document.getElementById('zalo-completion-banner');
                    if (banner) banner.style.display = 'block';

                    const detailsEl = document.getElementById('zalo-completion-details');
                    if (detailsEl && lastZaloDownloadedProject) {
                        detailsEl.textContent = `Đã lưu ${actualDownloaded} ảnh vào dự án ${lastZaloDownloadedProject}. ${warningText}`;
                    }

                    // Cập nhật thư viện ảnh ngay lập tức
                    if (typeof loadDailyPhotosStats === 'function') {
                        loadDailyPhotosStats();
                    }
                } else {
                    appendZaloLog('[Thất bại] Quá trình tải gặp sự cố.', 'error');
                    showToast('Tải ảnh Zalo không thành công!', 'error');
                }
            }
        }
    };

    // Try EventSource first
    try {
        zaloEventSource = new EventSource('/api/zalo/download/progress');
        zaloEventSource.onmessage = (event) => {
            try {
                const data = JSON.parse(event.data);
                onProgressData(data);
            } catch (e) {
                console.error('SSE parse error:', e);
            }
        };
        zaloEventSource.onerror = () => {
            if (zaloEventSource) {
                zaloEventSource.close();
                zaloEventSource = null;
            }
            if (!fallbackPollTimer && !isDone) {
                fallbackPollTimer = setInterval(async () => {
                    try {
                        const res = await fetch('/api/zalo/download/progress/poll');
                        const pData = await res.json();
                        if (pData.success && pData.data) {
                            onProgressData(pData.data);
                        }
                    } catch {}
                }, 1000);
            }
        };
    } catch (e) {
        fallbackPollTimer = setInterval(async () => {
            try {
                const res = await fetch('/api/zalo/download/progress/poll');
                const pData = await res.json();
                if (pData.success && pData.data) {
                    onProgressData(pData.data);
                }
            } catch {}
        }, 1000);
    }
}

// Update progress elements
function updateZaloProgressUI(p) {
    const statTotal = document.getElementById('zalo-stat-total');
    const statDownloaded = document.getElementById('zalo-stat-downloaded');
    const statSkipped = document.getElementById('zalo-stat-skipped');
    const statFailed = document.getElementById('zalo-stat-failed');
    const statLowQuality = document.getElementById('zalo-stat-low-quality');
    const statUnknownDate = document.getElementById('zalo-stat-unknown-date');
    const fill = document.getElementById('zalo-progress-fill');
    const statusText = document.getElementById('zalo-progress-status');
    const percentText = document.getElementById('zalo-progress-percent');
    const currentFileText = document.getElementById('zalo-current-file-text');

    const total = p.total || 0;
    const downloaded = p.downloaded || 0;
    const skipped = p.skipped || 0;
    const failed = p.failed || 0;
    const actualNew = p.counterVersion === 2 ? downloaded : Math.max(0, downloaded - skipped);

    if (statTotal) statTotal.textContent = total;
    if (statDownloaded) statDownloaded.textContent = actualNew;
    if (statSkipped) statSkipped.textContent = skipped;
    if (statFailed) statFailed.textContent = failed;
    if (statLowQuality) statLowQuality.textContent = p.lowQuality || 0;
    if (statUnknownDate) statUnknownDate.textContent = p.unknownDate || 0;

    let percent = 0;
    if (total > 0) {
        const processed = p.counterVersion === 2 ? (p.current || 0) : downloaded + failed;
        percent = Math.min(100, Math.round((processed / total) * 100));
    } else if (p.status === 'done') {
        percent = 100;
    }

    if (fill) {
        fill.style.width = percent + '%';
        if (p.status === 'downloading') {
            fill.classList.add('animated');
        } else {
            fill.classList.remove('animated');
        }
    }

    if (percentText) percentText.textContent = percent + '%';

    if (statusText) {
        if (p.status === 'downloading') {
            statusText.textContent = `Đang tải: ${downloaded}/${total} ảnh...`;
        } else if (p.status === 'done') {
            statusText.textContent = `Hoàn thành (${total} ảnh)`;
        } else if (p.status === 'error') {
            statusText.textContent = 'Gặp lỗi trong quá trình tải';
        } else {
            statusText.textContent = 'Sẵn sàng tải ảnh...';
        }
    }

    if (currentFileText) {
        if (p.status === 'done') {
            currentFileText.textContent = `Đã lưu ${actualNew} ảnh; ${p.lowQuality || 0} ảnh xem trước; ${p.unknownDate || 0} ảnh không rõ ngày (lưu vào unknown-date); ${failed} lỗi tải.`;
        } else if (p.currentFile) {
            currentFileText.textContent = `File hiện tại: ${p.currentFile}`;
        }
    }
}

// Append log to terminal
function appendZaloLog(msg, type = 'info') {
    const term = document.getElementById('zalo-log-terminal');
    if (!term) return;

    const timeStr = new Date().toLocaleTimeString('vi-VN');
    const line = document.createElement('div');
    line.className = `log-line ${type}`;
    line.textContent = `[${timeStr}] ${msg}`;

    term.appendChild(line);
    term.scrollTop = term.scrollHeight;
}

function clearZaloLogs() {
    const term = document.getElementById('zalo-log-terminal');
    if (term) term.innerHTML = '';
}

// Mở thư mục input_images trong Windows Explorer
async function openInputImagesFolder() {
    try {
        const res = await fetch('/api/open/folder', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ type: 'input_images' })
        });
        const data = await res.json();
        if (data.success) {
            showToast('Đã mở thư mục input_images trên máy', 'success');
        } else {
            showToast('Không thể mở thư mục: ' + data.error, 'error');
        }
    } catch (err) {
        showToast('Lỗi: ' + err.message, 'error');
    }
}

// Chuyển sang Tab 1 để xử lý
function navigateToProcessTab() {
    const projToUse = lastZaloDownloadedProject || currentProjectName;
    if (projToUse && typeof onProjectChange === 'function') {
        onProjectChange(projToUse);
    }
    switchTab('process', currentProcessSubtab, true);
}

// Chuyển sang Tab Quản Lý Ảnh (Thư Viện Ảnh)
async function navigateToPhotosTab(targetDay, targetProject) {
    const projToUse = targetProject || lastZaloDownloadedProject || currentProjectName;
    if (projToUse && typeof onProjectChange === 'function') {
        onProjectChange(projToUse);
    }
    switchTab('photos', 'daily', true);
    if (typeof loadDailyPhotosStats === 'function') {
        await loadDailyPhotosStats();
        if (targetDay) {
            selectDay(String(targetDay).padStart(2, '0'));
        }
    }
}

// Tự động kiểm tra trạng thái Zalo khi mở trang web
document.addEventListener('DOMContentLoaded', () => {
    setTimeout(() => {
        checkZaloStatus();
    }, 1000);
});

// Đóng dropdown khi nhấp chuột ra ngoài
document.addEventListener('click', (e) => {
    const picker = document.getElementById('zalo-group-picker-wrapper');
    const menu = document.getElementById('zalo-group-dropdown-menu');
    if (picker && menu && menu.style.display === 'block') {
        if (!picker.contains(e.target)) {
            toggleZaloGroupDropdown(false);
        }
    }
});


// ==================== BỔ SUNG ẢNH CHẤM CÔNG (SUPPLEMENT) ====================


// Keep destructive actions pending through confirmation, request, and refresh.
(() => {
    const pending = new Set();
    function guard(name, action) {
        return async function(...args) {
            const key = name + ':' + JSON.stringify(args);
            if (pending.has(key)) return;
            pending.add(key);
            const button = document.activeElement;
            const isButton = button && button.tagName === 'BUTTON';
            if (isButton) button.disabled = true;
            try {
                return await action.apply(this, args);
            } finally {
                pending.delete(key);
                if (isButton && button.isConnected) {
                    button.disabled = false;
                    if (document.activeElement === document.body) button.focus();
                }
            }
        };
    }
    confirmClearFaceCache = guard('confirmClearFaceCache', confirmClearFaceCache);
    deleteExcelUpload = guard('deleteExcelUpload', deleteExcelUpload);
    deleteExcelExtractedFolder = guard('deleteExcelExtractedFolder', deleteExcelExtractedFolder);
    deleteExcelExtractedFile = guard('deleteExcelExtractedFile', deleteExcelExtractedFile);
    deleteExcelFaceFolder = guard('deleteExcelFaceFolder', deleteExcelFaceFolder);
    deleteExcelFaceFile = guard('deleteExcelFaceFile', deleteExcelFaceFile);
    deletePDFUpload = guard('deletePDFUpload', deletePDFUpload);
    deletePDFExtractedFolder = guard('deletePDFExtractedFolder', deletePDFExtractedFolder);
    deletePDFExtractedFile = guard('deletePDFExtractedFile', deletePDFExtractedFile);
    deletePDFFaceFolder = guard('deletePDFFaceFolder', deletePDFFaceFolder);
    deletePDFFaceFile = guard('deletePDFFaceFile', deletePDFFaceFile);
    deleteResultFile = guard('deleteResultFile', deleteResultFile);
    confirmDeleteProject = guard('confirmDeleteProject', confirmDeleteProject);
    deleteDailyPhoto = guard('deleteDailyPhoto', deleteDailyPhoto);
    confirmDeleteAllDayPhotos = guard('confirmDeleteAllDayPhotos', confirmDeleteAllDayPhotos);
    submitDeleteAllDaysPhotosDirect = guard('submitDeleteAllDaysPhotosDirect', submitDeleteAllDaysPhotosDirect);
    deleteEmployeePhoto = guard('deleteEmployeePhoto', deleteEmployeePhoto);
    deleteEmployee = guard('deleteEmployee', deleteEmployee);
    supplementProcess = guard('supplementProcess', supplementProcess);
    supplementProcessAll = guard('supplementProcessAll', supplementProcessAll);
    supplementDelete = guard('supplementDelete', supplementDelete);
})();

async function supplementLoadEmployees() {
    const select = document.getElementById('supp-employee');
    if (!select) return;
    const projectId = allProjectsList.find(p => p.name === currentProjectName)?.project_id;
    if (!projectId) return;
    try {
        const data = await apiGet(`/api/portraits?include_history=true&project_id=${encodeURIComponent(projectId)}`);
        if (!data.success) throw new Error(data.error || 'Employee loading failed');
        select.replaceChildren();
        for (const employee of data.employees || []) {
            const option = document.createElement('option'); option.value = employee.employee_id; option.textContent = employee.name; select.appendChild(option);
        }
    } catch (error) { showToast(error.message, 'error'); }
}

function supplementFileSelected(input) {
    const nameEl = document.getElementById('supp-file-name');
    if (input.files && input.files[0]) {
        nameEl.textContent = input.files[0].name;
    } else {
        nameEl.textContent = '';
    }
}

function supplementDropFile(e) {
    e.preventDefault();
    const zone = document.getElementById('supplement-upload-zone');
    zone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files && files.length > 0) {
        const input = document.getElementById('supp-file');
        input.files = files;
        supplementFileSelected(input);
    }
}

async function supplementUpload(e) {
    e.preventDefault();
    const form = document.getElementById('supplement-upload-form');
    const fileInput = document.getElementById('supp-file');
    if (!fileInput.files || fileInput.files.length === 0) {
        alert('Vui lòng chọn file ảnh!');
        return;
    }
    
    const btn = document.getElementById('btn-supp-upload');
    btn.disabled = true;
    btn.innerHTML = 'Đang upload...';
    
    try {
        const formData = new FormData();
        formData.append('photo', fileInput.files[0]);
        formData.append('employee_id', document.getElementById('supp-employee').value);
        formData.append('project_id', allProjectsList.find(p => p.name === currentProjectName)?.project_id || '');
        formData.append('target_date', document.getElementById('supp-date').value);
        const targetTime = document.getElementById('supp-time').value.trim();
        if (targetTime) formData.append('target_time', targetTime);
        formData.append('watermark_style', document.getElementById('supp-style').value);
        formData.append('watermark_position', document.getElementById('supp-position').value);
        formData.append('location_name', document.getElementById('supp-location').value);
        formData.append('gps_coords', document.getElementById('supp-gps').value);
        formData.append('remove_old_watermark', targetTime && document.getElementById('supp-remove-old').checked ? 'true' : 'false');
        formData.append('modify_exif', targetTime && document.getElementById('supp-modify-exif').checked ? 'true' : 'false');
        
        const resp = await fetch('/api/supplement/upload', { method: 'POST', body: formData });
        const data = await resp.json();
        
        if (resp.ok && data.success) {
            alert('✅ Upload thành công! Record ID: ' + data.record.id);
            fileInput.value = '';
            document.getElementById('supp-file-name').textContent = '';
            supplementLoadRecords();
        } else {
            alert('❌ Lỗi: ' + (data.error || 'Không rõ'));
        }
    } catch (err) {
        alert('❌ Lỗi kết nối: ' + err.message);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg> Upload & Tạo Record';
    }
}

async function supplementLoadRecords() {
    const container = document.getElementById('supplement-records-container');
    if (!container) return;
    container.innerHTML = '<p style="text-align:center; padding:16px; color:var(--text-secondary)">Đang tải...</p>';
    
    try {
        const resp = await fetch(`/api/supplement/records?project_id=${encodeURIComponent(allProjectsList.find(p => p.name === currentProjectName)?.project_id || '')}`);
        const data = await resp.json();
        
        if (!data.success || !data.records || data.records.length === 0) {
            container.innerHTML = '<p style="color:var(--text-secondary); text-align:center; padding:24px;">Chưa có ảnh bổ sung nào.</p>';
            return;
        }
        
        let html = '<table class="data-table" style="width:100%;">';
        html += '<thead><tr>';
        html += '<th>ID</th><th>Nhân viên</th><th>Ngày</th><th>Giờ</th><th>Trạng thái</th><th>Hành động</th>';
        html += '</tr></thead><tbody>';
        
        for (const r of data.records) {
            const statusBadge = {
                'pending': '<span style="color:#f59e0b;">&#9679; Chờ xử lý</span>',
                'processing': '<span style="color:#3b82f6;">&#9679; Đang xử lý</span>',
                'done': '<span style="color:#22c55e;">&#9679; Hoàn thành</span>',
                'error': '<span style="color:#ef4444;">&#9679; Lỗi</span>',
            }[r.status] || r.status;
            
            let actions = '';
            if (r.status === 'pending') {
                actions += `<button class="btn btn-primary btn-sm" onclick="supplementProcess('${r.id}')">Xử lý</button> `;
            }
            if (r.status === 'done') {
                actions += `<button class="btn btn-success btn-sm" onclick="supplementDownload('${r.id}')">Tải</button> `;
                actions += `<button class="btn btn-secondary btn-sm" onclick="supplementPreview('${r.id}')">Xem</button> `;
            }
            if (r.status === 'pending' || r.status === 'error') {
                actions += `<button class="btn btn-secondary btn-sm" onclick="supplementPreview('${r.id}')">Preview</button> `;
            }
            actions += `<button class="btn btn-sm" style="color:#ef4444;" onclick="supplementDelete('${r.id}')">Xóa</button>`;
            
            const dateFormatted = r.target_date ? r.target_date.split('-').reverse().join('/') : '';
            
            html += `<tr>`;
            html += `<td><code>${r.id}</code></td>`;
            html += `<td>${r.employee_name || ''}</td>`;
            html += `<td>${dateFormatted}</td>`;
            html += `<td>${r.target_time || ''}</td>`;
            html += `<td>${statusBadge}`;
            for (const [label,value] of [['Date',r.date_status],['Face',r.face_status],['Integrity',r.integrity_status],['Storage',r.storage_status],['Watermark',r.processing?.watermark],['EXIF',r.processing?.exif],['Can apply',r.can_apply],['Reasons',r.reasons]]) {
                const text = `${label}: ${typeof value === 'object' ? JSON.stringify(value) : value ?? 'unknown'}`;
                html += `<div>${text.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;')}</div>`;
            }
            html += '</td>';
            html += `<td style="white-space:nowrap;">${actions}</td>`;
            html += `</tr>`;
            
            if (r.status === 'error' && r.error_message) {
                html += `<tr><td colspan="6" style="color:#ef4444; font-size:0.85rem; padding:4px 12px;">❌ ${r.error_message}</td></tr>`;
            }
        }
        
        html += '</tbody></table>';
        container.innerHTML = html;
    } catch (err) {
        container.innerHTML = `<p style="color:#ef4444; text-align:center; padding:16px;">Lỗi: ${err.message}</p>`;
    }
}

function supplementGenerationFailed(data) {
    const failed = /fail|error|missing|needs_confirmation/i;
    const records = data.records || [data.record || data];
    return records.some(record => [record.status, record.watermark_status, record.processing?.watermark, record.processing?.exif]
        .some(value => failed.test(typeof value === 'object' ? value?.status || '' : value || '')));
}

async function supplementProcess(recordId) {
    if (!await confirmAction('Bạn có chắc muốn xử lý record này?')) return;
    try {
        const resp = await fetch('/api/supplement/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ record_id: recordId })
        });
        const data = await resp.json();
        if (resp.ok && data.success && !supplementGenerationFailed(data)) {
            alert('✅ Xử lý thành công!');
        } else {
            alert('❌ Lỗi: ' + (data.error || 'Không rõ'));
        }
        supplementLoadRecords();
    } catch (err) {
        alert('❌ Lỗi: ' + err.message);
    }
}

async function supplementProcessAll() {
    if (!await confirmAction('Xử lý tất cả ảnh đang chờ?')) return;
    try {
        const resp = await fetch('/api/supplement/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({})
        });
        const data = await resp.json();
        if (resp.ok && data.success && !supplementGenerationFailed(data)) {
            const count = data.records ? data.records.length : 0;
            alert(`✅ Đã xử lý ${count} ảnh!`);
        } else {
            alert('❌ Lỗi: ' + (data.error || 'Không rõ'));
        }
        supplementLoadRecords();
    } catch (err) {
        alert('❌ Lỗi: ' + err.message);
    }
}

function supplementDownload(recordId) {
    window.open(`/api/supplement/download/${recordId}`, '_blank');
}

async function supplementPreview(recordId) {
    const modal = document.getElementById('supplement-preview-modal');
    const img = document.getElementById('supp-preview-img');
    const title = document.getElementById('supp-preview-title');
    const validationDiv = document.getElementById('supp-preview-validation');
    
    img.src = `/api/supplement/preview/${recordId}?t=${Date.now()}`;
    title.textContent = `Preview - ${recordId}`;
    validationDiv.innerHTML = 'Đang kiểm tra chất lượng...';
    modal.style.display = 'flex';
    
    // Load validation
    try {
        const resp = await fetch(`/api/supplement/validate/${recordId}`);
        const data = await resp.json();
        if (data.success && data.validation) {
            const v = data.validation;
            let html = '<div style="padding:8px; background:var(--bg-secondary); border-radius:8px; font-size:0.88rem;">';
            html += `<p style="margin:0 0 6px;"><strong>Kết quả kiểm tra:</strong> ${v.valid ? '✅ Đạt' : '❌ Không đạt'}</p>`;
            if (v.checks) {
                for (const [key, check] of Object.entries(v.checks)) {
                    html += `<p style="margin:2px 0;">${check.ok ? '✅' : '❌'} ${key}: ${JSON.stringify(check)}</p>`;
                }
            }
            if (v.warnings && v.warnings.length > 0) {
                html += '<p style="margin:6px 0 0; color:#f59e0b;">⚠️ ' + v.warnings.join(' | ') + '</p>';
            }
            html += '</div>';
            validationDiv.innerHTML = html;
        } else {
            validationDiv.innerHTML = '';
        }
    } catch {
        validationDiv.innerHTML = '';
    }
}

async function supplementDelete(recordId) {
    if (!await confirmAction('Xóa record ' + recordId + '? File ảnh liên quan cũng sẽ bị xóa.')) return;
    try {
        const resp = await fetch('/api/supplement/delete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ record_id: recordId })
        });
        const data = await resp.json();
        if (resp.ok && data.success) {
            supplementLoadRecords();
        } else {
            alert('❌ Lỗi xóa: ' + (data.error || 'Không rõ'));
        }
    } catch (err) {
        alert('❌ Lỗi: ' + err.message);
    }
}

async function supplementLoadMissing() {
    const card = document.getElementById('supplement-missing-card');
    const list = document.getElementById('supplement-missing-list');
    card.style.display = 'block';
    list.innerHTML = '<p style="text-align:center; color:var(--text-secondary);">Đang quét file chấm công...</p>';
    
    try {
        const resp = await fetch('/api/supplement/missing');
        const data = await resp.json();
        
        if (!data.success || !data.missing || data.missing.length === 0) {
            list.innerHTML = '<p style="text-align:center; color:#22c55e; padding:12px;">✅ Không có ngày thiếu ảnh nào!</p>';
            return;
        }
        
        let html = '<table class="data-table" style="width:100%; font-size:0.88rem;">';
        html += '<thead><tr><th>Nhân viên</th><th>Ngày</th><th>Thứ</th><th>Vấn đề</th><th></th></tr></thead><tbody>';
        
        for (const m of data.missing) {
            html += '<tr>';
            html += `<td>${m.person_name || ''}</td>`;
            html += `<td>${m.date || ''}</td>`;
            html += `<td>${m.weekday || ''}</td>`;
            html += `<td style="color:#f59e0b;">${m.issue_description || ''}</td>`;
            html += `<td><button class="btn btn-primary btn-sm" onclick="supplementFillFromMissing('${m.employee_id || ''}','${m.date}')">Bổ sung</button></td>`;
            html += '</tr>';
        }
        
        html += '</tbody></table>';
        html += `<p style="margin-top:8px; font-size:0.85rem; color:var(--text-secondary);">Tổng: ${data.missing.length} bản ghi thiếu</p>`;
        list.innerHTML = html;
    } catch (err) {
        list.innerHTML = `<p style="color:#ef4444;">❌ Lỗi: ${err.message}</p>`;
    }
}

function supplementFillFromMissing(name, date) {
    // Convert dd/mm/yyyy to yyyy-mm-dd for the date input
    let isoDate = date;
    if (date && date.includes('/')) {
        const parts = date.split('/');
        if (parts.length === 3) {
            isoDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
        }
    }
    const employee = allEmployeesList.find(employee => employee.employee_id === name);
    document.getElementById('supp-employee').value = employee?.employee_id || '';
    setDatePickerValue(document.getElementById('supp-date'), isoDate);
    document.getElementById('supp-time').value = '';
    // Scroll to upload form
    document.getElementById('supplement-upload-form').scrollIntoView({ behavior: 'smooth' });
}

async function confirmEmployeeIdentity(employee_id) {
    const payroll_code = prompt('Mã chấm công chính xác (giữ số 0 đầu):');
    if (!payroll_code) return;
    const valid_from = prompt('Ngày bắt đầu thuộc dự án (YYYY-MM-DD), theo hồ sơ:');
    if (!valid_from) return;
    const reviewer = 'system';
    try {
        const res = await apiPost('/api/portraits/employee/confirm', {project_id: allProjectsList.find(p => p.name === currentProjectName)?.project_id, project: currentProjectName, employee_id, payroll_code, valid_from, reviewer});
        if (!res.success) throw new Error(res.error);
        await loadPortraits();
        showToast('Đã xác nhận mã chấm công', 'success');
    } catch (err) { showToast(err.message, 'error'); }
}

async function generateInternalEmployeeCodes(button) {
    if (button && button.disabled) return;
    if (button) button.disabled = true;
    try {
        const res = await apiPost('/api/portraits/internal-codes/generate', {});
        if (!res.success) throw new Error(res.error || 'Không tạo được mã nội bộ');
        await loadPortraits();
        showToast('Đã tạo ' + res.created_count + ' mã nội bộ; giữ nguyên ' +
            res.unchanged_count + ' mã đã có trên tất cả dự án.', 'success');
    } catch (err) {
        showToast(err.message, 'error');
    } finally {
        if (button) button.disabled = false;
    }
}
