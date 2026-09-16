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

async function apiGet(url) {
    const response = await fetch(url);
    return response.json();
}

async function apiPost(url, data = {}) {
    const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
    });
    return response.json();
}

function initializeAggregateReportSettings() {
    const fromInput = document.getElementById('report-from-date');
    const toInput = document.getElementById('report-to-date');
    if (!fromInput || !toInput) return;

    const now = new Date();
    const localToday = new Date(now.getTime() - now.getTimezoneOffset() * 60000)
        .toISOString().slice(0, 10);
    const firstDay = `${localToday.slice(0, 8)}01`;
    if (!fromInput.value) fromInput.value = firstDay;
    if (!toInput.value) toInput.value = localToday;
}

function getAggregateReportSettings(projectFallback) {
    const projectInput = document.getElementById('project-name');
    const fromInput = document.getElementById('report-from-date');
    const toInput = document.getElementById('report-to-date');
    const reportProjectName = projectInput && projectInput.value.trim()
        ? projectInput.value.trim()
        : projectFallback;
    const fromDate = fromInput ? fromInput.value : '';
    const toDate = toInput ? toInput.value : '';

    if (!reportProjectName) {
        showToast('Vui lòng nhập tên dự án cho báo cáo tổng hợp', 'warning');
        return null;
    }
    if (!fromDate || !toDate) {
        showToast('Vui lòng chọn đầy đủ Từ ngày và Đến ngày', 'warning');
        return null;
    }
    if (fromDate > toDate) {
        showToast('Từ ngày không được sau Đến ngày', 'warning');
        return null;
    }
    return {
        report_project_name: reportProjectName,
        from_date: fromDate,
        to_date: toDate
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
    if (!VALID_TABS[tabId]) {
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
        if (typeof initZaloTab === 'function') {
            initZaloTab();
        }
    } else if (tabId === 'supplement') {
        if (typeof supplementLoadRecords === 'function') {
            supplementLoadRecords();
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

    if (tabParam && VALID_TABS[tabParam]) {
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
    initializeAggregateReportSettings();
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
        } else {
            showToast(result.error || 'Lỗi upload', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

async function extractPDF() {
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
        const result = await apiPost('/api/pdf/extract', { filename: pdfFilename });

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
    showToast(`Đã chọn: ${filename}`, 'info');
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
                                    <h3 style="margin:0;">${folder.folder} (${folder.count} files)</h3>
                                    <button class="btn btn-primary btn-sm" onclick="startPDFFaceAnalyze('${safeFolder}')">Quét mặt</button>
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

async function startPDFFaceAnalyze(folder) {
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
            showToast('Đã bắt đầu quét mặt từ PDF...', 'info');
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
        document.getElementById('pdf-face-progress-title').textContent =
            result.current ? `Đang xuất: ${result.current}` : 'Đang quét mặt từ PDF...';
        document.getElementById('pdf-face-progress-percent').textContent = `${pct}%`;
        document.getElementById('pdf-face-progress-fill').style.width = `${pct}%`;
        document.getElementById('pdf-face-progress-detail').textContent =
            `${result.progress}/${result.total} file`;

        if (result.status === 'completed') {
            showToast(`Hoàn thành! Đã tạo file Word theo nhân viên và báo cáo giải trình tổng hợp.`, 'success', 'Xem Kết Quả Ngay', () => navigateToResults('summary'));
            loadPDFFaceFiles();
            loadAggregateReports();
            setTimeout(() => {
                document.getElementById('pdf-face-progress-section').style.display = 'none';
            }, 3000);
        } else if (result.status === 'failed') {
            showToast((result.errors && result.errors[0]) || 'Lỗi quét mặt từ PDF', 'error');
        } else {
            setTimeout(checkPDFFaceProgress, 1000);
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
        } else {
            showToast(result.error || 'Lỗi upload', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

async function extractExcel() {
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
        const result = await apiPost('/api/excel/extract', { filename: excelFilename });
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
    showToast(`Đã chọn: ${filename}`, 'info');
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
                                    <h3 style="margin:0;">${folder.folder} (${folder.count} người)</h3>
                                    <button class="btn btn-primary btn-sm" onclick="startExcelFaceAnalyze('${safeFolder}')">Quét mặt</button>
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

async function startExcelFaceAnalyze(folder) {
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
            showToast('Đã bắt đầu quét mặt...', 'info');
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
        document.getElementById('excel-face-progress-title').textContent =
            result.current ? `Đang xuất: ${result.current}` : 'Đang quét mặt...';
        document.getElementById('excel-face-progress-percent').textContent = `${pct}%`;
        document.getElementById('excel-face-progress-fill').style.width = `${pct}%`;
        document.getElementById('excel-face-progress-detail').textContent =
            `${result.progress}/${result.total} file`;

        if (result.status === 'completed') {
            showToast(`Hoàn thành! Đã tạo file Word theo nhân viên và báo cáo giải trình tổng hợp.`, 'success', 'Xem Kết Quả Ngay', () => navigateToResults('summary'));
            loadExcelFaceFiles();
            loadAggregateReports();
            setTimeout(() => {
                document.getElementById('excel-face-progress-section').style.display = 'none';
            }, 3000);
        } else if (result.status === 'failed') {
            showToast('Lỗi: ' + (result.errors[0] || 'Không rõ'), 'error');
        } else {
            setTimeout(checkExcelFaceProgress, 1000);
        }
    } catch (error) {
        console.error('Lỗi kiểm tra tiến độ quét mặt:', error);
        setTimeout(checkExcelFaceProgress, 2000);
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file "${filename}"?\nHành động này không thể hoàn tác.`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa toàn bộ đợt "${folder}"?\nTất cả file Word trong đợt này sẽ bị xóa.`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa toàn bộ kết quả quét mặt "${folder}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file PDF "${filename}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa toàn bộ đợt tách PDF "${folder}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa toàn bộ kết quả quét mặt "${folder}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file "${filename}"?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn xóa file kết quả "${filename}"?`)) return;
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
    if (!confirm('Lưu trữ dự án này? Giữ nguyên ảnh và lịch sử.')) return;
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
        input.value = now.getFullYear() + '-' + String(now.getMonth()+1).padStart(2, '0');
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
    const reviewer = prompt('Tên người xác nhận:');
    if (!reviewer) return;
    if (!confirm('Chỉ xác nhận nếu mọi ảnh trong thư mục thuộc cùng ngày. Nếu lẫn tháng, hãy phân loại từng ảnh trước.')) return;
    const result = await apiPost('/api/photos/daily/legacy-map', {
        project: currentProjectName, folder, date: target, reviewer, confirm_single_period: true
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

            gallery.innerHTML = res.photos.map(p => `
                <div class="thumb-card">
                    <div class="thumb-img-wrap" onclick="openLightbox('${p.url}', 'Ảnh camera ngày ${day} - ${p.filename}')">
                        <img src="${p.url}" alt="${p.filename}" loading="lazy">
                    </div>
                    <div class="thumb-info">
                        <div class="thumb-name" title="${p.filename}">${p.filename}</div>
                        <div class="thumb-size">${p.size_kb} KB</div>
                    </div>
                    <div class="thumb-actions">
                        <button class="btn btn-icon-danger btn-sm" onclick="deleteDailyPhoto('${p.filename}')" title="Xóa ảnh này">
                            ${TRASH_ICON_SVG}
                        </button>
                    </div>
                </div>
            `).join('');
        }
    } catch (err) {
        console.error('Lỗi nạp ảnh ngày:', err);
    }
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
    if (!confirm(`Xóa ảnh camera "${filename}" của ngày ${activeSelectedDay}?`)) return;
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
    if (!confirm(`Bạn có chắc chắn muốn XÓA TOÀN BỘ ảnh camera của ngày ${activeSelectedDay} không?`)) return;
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

function openCreateEmployeeModal() {
    const nameInput = document.getElementById('new-employee-name');
    const fileInput = document.getElementById('new-employee-photos');
    if (nameInput) nameInput.value = '';
    if (fileInput) fileInput.value = '';
    openModal('modal-create-employee');
    if (nameInput) nameInput.focus();
}

async function submitCreateEmployee() {
    const nameInput = document.getElementById('new-employee-name');
    const fileInput = document.getElementById('new-employee-photos');
    const name = nameInput ? nameInput.value.trim() : '';

    if (!name) {
        showToast('Vui lòng nhập họ tên nhân viên', 'warning');
        return;
    }

    const payroll_code = prompt('Mã chấm công chính xác (giữ số 0 đầu):');
    if (!payroll_code) return;
    const valid_from = prompt('Ngày bắt đầu thuộc dự án (YYYY-MM-DD):');
    if (!valid_from) return;
    const reviewer = prompt('Người xác nhận:');
    if (!reviewer) return;
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

    const effective_date = prompt('Ngày chuyển có hiệu lực (YYYY-MM-DD):');
    if (!effective_date) return;
    const payroll_code = prompt('Mã chấm công tại dự án đích:');
    if (!payroll_code) return;
    const reviewer = prompt('Người xác nhận:');
    if (!reviewer) return;
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
    openModal('modal-employee-photos');
}

function refreshEmployeePhotosModal() {
    const gallery = document.getElementById('emp-photos-gallery');
    if (!gallery) return;

    const emp = allEmployeesList.find(e => e.employee_id === activeEmpModalName);
    if (!emp || !emp.images || emp.images.length === 0) {
        gallery.innerHTML = `
            <div style="grid-column: 1 / -1; padding: 18px; text-align: center; color: var(--text-muted); background: var(--bg-subtle); border-radius: var(--radius);">
                Chưa có ảnh chân dung nào cho nhân viên này. Nhấp ô phía trên để tải ảnh lên.
            </div>
        `;
        return;
    }

    const safeProj = encodeURIComponent(currentProjectName);
    const safeName = encodeURIComponent(activeEmpModalName);

    gallery.innerHTML = emp.images.map(img => {
        const url = emp.image_urls[img];
        const safeImg = img.replace(/'/g, "\\'");
        return `
            <div class="thumb-card">
                <div class="thumb-img-wrap" onclick="openLightbox('${url}', '${emp.name} - ${img}')">
                    <img src="${url}" alt="${img}" loading="lazy">
                </div>
                <div class="thumb-info">
                    <div class="thumb-name" title="${img}">${img}</div>
                </div>
                <div class="thumb-actions">
                    <button class="btn btn-icon-danger btn-sm" onclick="deleteEmployeePhoto('${safeImg}')" title="Xóa ảnh này">
                        ${TRASH_ICON_SVG}
                    </button>
                </div>
            </div>
        `;
    }).join('');
}

async function handleUploadEmployeePhoto(input) {
    if (!input.files || input.files.length === 0 || !activeEmpModalName) return;
    const formData = new FormData();
    formData.append('project', currentProjectName);
            formData.append('project_id', allProjectsList.find(p => p.name === currentProjectName)?.project_id || '');
    formData.append('employee_id', activeEmpModalName);
    const reviewer = prompt('Người xác nhận ảnh chân dung:');
    if (!reviewer) return;
    formData.append('reviewer', reviewer);
    for (let i = 0; i < input.files.length; i++) {
        formData.append('files', input.files[i]);
    }

    try {
        const res = await fetch('/api/portraits/employee/upload', {
            method: 'POST',
            body: formData
        }).then(r => r.json());

        if (res.success) {
            showToast(`Đã thêm ${res.saved_count} ảnh chân dung cho ${activeEmpModalName}`, 'success');
            input.value = '';
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
    if (!confirm(`Bạn có chắc muốn xóa ảnh chân dung "${filename}" của ${activeEmpModalName}?`)) return;
    try {
        const res = await apiPost('/api/portraits/employee/delete-photo', {
            project_id: allProjectsList.find(p => p.name === currentProjectName)?.project_id,
            project: currentProjectName,
            employee_id: activeEmpModalName,
            filename: filename
        });
        if (res.success) {
            showToast('Đã xóa ảnh chân dung', 'success');
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

async function deleteEmployee(empName) {
    if (!confirm(`Bạn có chắc chắn muốn xóa nhân viên "${empName}" khỏi danh sách hoạt động? Ảnh và lịch sử vẫn được giữ.`)) return;
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

// ==================== TAB: TẢI ẢNH ZALO ====================

let zaloPollTimer = null;
let zaloEventSource = null;
let zaloGroupsList = [];
let zaloIsDownloading = false;

function initZaloTab() {
    setZaloDatePreset('today');
    loadProjectsForZalo();
    checkZaloStatus();

    // Re-render cached groups & restore selected group UI when switching tabs
    if (zaloGroupsList.length > 0) {
        renderZaloGroupItems(zaloGroupsList);
        restoreZaloGroupSelection();
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
        if (subtitle) subtitle.textContent = 'Dịch vụ Zalo Service chưa phản hồi. Vui lòng kiểm tra lại dịch vụ.';
        if (showToastMsg) showToast('Lỗi kết nối Zalo Service: ' + err.message, 'error');
    }
}

// Bắt đầu đăng nhập bằng QR
async function startZaloQrLogin(force = false) {
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

function cancelZaloLogin() {
    if (zaloPollTimer) {
        clearInterval(zaloPollTimer);
        zaloPollTimer = null;
    }
    const qrCard = document.getElementById('zalo-qr-card');
    if (qrCard) qrCard.style.display = 'none';
    appendZaloLog('[Hệ thống] Đã đóng cửa sổ đăng nhập QR.', 'info');
}

async function logoutZalo() {
    if (!confirm('Bạn có chắc chắn muốn đăng xuất tài khoản Zalo không?')) return;
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

        item.innerHTML = `
            <div class="zalo-item-name">
                <span class="zalo-item-icon">${escapeHtml(firstLetter)}</span>
                <span>${escapeHtml(g.name)}</span>
            </div>
            <span class="zalo-item-count">${g.totalMember || 0} thành viên</span>
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
        fromInput.value = todayStr;
        toInput.value = todayStr;
    } else if (preset === 'yesterday') {
        const y = new Date();
        y.setDate(y.getDate() - 1);
        const yStr = formatDate(y);
        fromInput.value = yStr;
        toInput.value = yStr;
    } else if (preset === '3days') {
        const d = new Date();
        d.setDate(d.getDate() - 2);
        fromInput.value = formatDate(d);
        toInput.value = todayStr;
    } else if (preset === '7days') {
        const d = new Date();
        d.setDate(d.getDate() - 6);
        fromInput.value = formatDate(d);
        toInput.value = todayStr;
    } else if (preset === 'all') {
        fromInput.value = '';
        toInput.value = '';
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
    const msgLimit = parseInt(document.getElementById('zalo-msg-limit')?.value || '200', 10);

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
            throw new Error(data.error || 'Lỗi bắt đầu tải');
        }

        appendZaloLog(`[Zalo] ${data.message || 'Đã tìm thấy tin nhắn, đang tải ảnh...'}`, 'info');

        // Lắng nghe tiến trình tải
        listenZaloProgress();

    } catch (err) {
        showToast('Lỗi tải ảnh: ' + err.message, 'error');
        appendZaloLog('[Lỗi] ' + err.message, 'error');
        zaloIsDownloading = false;
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
                qrMsg.innerHTML = '<span class="status-indicator-dot warning"></span> <strong>Quét mã QR dưới đây để liên kết Chrome tải ảnh</strong> (chỉ cần làm 1 lần)...';
            }
        } else if (p.status && p.status !== 'waiting_qr' && p.status !== 'idle' && !p.qrImage) {
            if (qrCard && qrCard.style.display !== 'none') {
                qrCard.style.display = 'none';
            }
        }

        if (p.errors && p.errors.length > 0) {
            const lastErr = p.errors[p.errors.length - 1];
            appendZaloLog(`[Lỗi file] ${lastErr}`, 'error');
        }

        if (p.status === 'done' || p.status === 'error') {
            if (!isDone) {
                isDone = true;
                zaloIsDownloading = false;
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
                    const warningText = `Ảnh xem trước: ${p.lowQuality || 0}; bỏ qua vì không rõ ngày gửi: ${p.unknownDate || 0}; lỗi tải: ${p.failed || 0}.`;
                    const hasWarnings = (p.lowQuality || 0) + (p.unknownDate || 0) + (p.failed || 0) > 0;
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
            currentFileText.textContent = `Đã lưu ${actualNew} ảnh; ${p.lowQuality || 0} ảnh xem trước; ${p.unknownDate || 0} ảnh không rõ ngày gửi; ${failed} lỗi tải.`;
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
        formData.append('employee_name', document.getElementById('supp-employee').value);
        formData.append('target_date', document.getElementById('supp-date').value);
        formData.append('target_time', document.getElementById('supp-time').value);
        formData.append('watermark_style', document.getElementById('supp-style').value);
        formData.append('watermark_position', document.getElementById('supp-position').value);
        formData.append('location_name', document.getElementById('supp-location').value);
        formData.append('gps_coords', document.getElementById('supp-gps').value);
        formData.append('remove_old_watermark', document.getElementById('supp-remove-old').checked ? 'true' : 'false');
        formData.append('modify_exif', document.getElementById('supp-modify-exif').checked ? 'true' : 'false');
        
        const resp = await fetch('/api/supplement/upload', { method: 'POST', body: formData });
        const data = await resp.json();
        
        if (data.success) {
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
        const resp = await fetch('/api/supplement/records');
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
            html += `<td>${statusBadge}</td>`;
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

async function supplementProcess(recordId) {
    if (!confirm('Bạn có chắc muốn xử lý record này?')) return;
    try {
        const resp = await fetch('/api/supplement/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ record_id: recordId })
        });
        const data = await resp.json();
        if (data.success) {
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
    if (!confirm('Xử lý tất cả ảnh đang chờ?')) return;
    try {
        const resp = await fetch('/api/supplement/process', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({})
        });
        const data = await resp.json();
        if (data.success) {
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
    if (!confirm('Xóa record ' + recordId + '? File ảnh liên quan cũng sẽ bị xóa.')) return;
    try {
        const resp = await fetch('/api/supplement/delete', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ record_id: recordId })
        });
        const data = await resp.json();
        if (data.success) {
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
            html += `<td><button class="btn btn-primary btn-sm" onclick="supplementFillFromMissing('${m.person_name}','${m.date}')">Bổ sung</button></td>`;
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
    document.getElementById('supp-employee').value = name;
    document.getElementById('supp-date').value = isoDate;
    document.getElementById('supp-time').value = '08:00';
    // Scroll to upload form
    document.getElementById('supplement-upload-form').scrollIntoView({ behavior: 'smooth' });
}

async function confirmEmployeeIdentity(employee_id) {
    const payroll_code = prompt('Mã chấm công chính xác (giữ số 0 đầu):');
    if (!payroll_code) return;
    const valid_from = prompt('Ngày bắt đầu thuộc dự án (YYYY-MM-DD), theo hồ sơ:');
    if (!valid_from) return;
    const reviewer = prompt('Họ tên người xác nhận danh tính:');
    if (!reviewer) return;
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
