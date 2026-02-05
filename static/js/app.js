/**
 * JavaScript cho phần mềm chấm công v2.0
 * Đơn giản hóa với 3 tab: Phân Tích + Tách PDF + Kết Quả
 */

// ==================== State ====================
let analysisResults = null;
let pdfTaskId = null;
let pdfFilename = null;
let logEventSource = null;

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

// ==================== Toast Notifications ====================

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;

    const icons = { success: '✅', error: '❌', warning: '⚠️', info: 'ℹ️' };
    toast.innerHTML = `<span>${icons[type]}</span><span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => toast.remove(), 4000);
}

// ==================== Tab Navigation ====================

document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();

        document.querySelectorAll('.nav-item').forEach(nav => nav.classList.remove('active'));
        item.classList.add('active');

        const tabId = item.dataset.tab;
        document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
        document.getElementById(`tab-${tabId}`).classList.add('active');

        const titles = {
            analyze: 'Phân Tích Chấm Công',
            pdf: 'Tách PDF Thành Word',
            results: 'Kết Quả'
        };
        document.querySelector('.page-title').textContent = titles[tabId] || 'Phân Tích';

        if (tabId === 'results') loadResultFiles();
        if (tabId === 'pdf') {
            loadPDFUploads();
            loadPDFExtractedFiles();
        }
    });
});

// ==================== Analysis ====================

async function runAnalysis() {
    const btn = document.getElementById('btn-analyze');
    const progressSection = document.getElementById('progress-section');
    const logPanel = document.getElementById('log-panel');

    btn.disabled = true;
    btn.textContent = '⏳ Đang xử lý...';
    progressSection.style.display = 'block';
    logPanel.style.display = 'block';
    clearLogs();

    updateProgress('Đang quét file chấm công...', 10);
    addLog('🚀 Bắt đầu phân tích...', 'info');

    // Connect to SSE for real-time logs
    startLogStream();

    try {
        // Step 1: Phân tích chấm công
        updateProgress('Đang phân tích ngày thiếu...', 30);
        const result = await apiPost('/api/analyze-full');

        // Stop SSE connection
        stopLogStream();

        if (result.success) {
            analysisResults = result;

            // Update stats
            document.getElementById('stat-files').textContent = result.summary.total_persons || 0;
            document.getElementById('stat-missing').textContent = result.summary.total_missing || 0;
            document.getElementById('stat-persons').textContent = result.summary.persons_with_issues || 0;
            document.getElementById('stat-matched').textContent = result.summary.total_matched || 0;

            // Update table
            updateProgress('Đang hiển thị kết quả...', 90);
            displayResults(result.records);

            updateProgress('Hoàn thành!', 100);
            addLog(`✅ Hoàn thành! Tìm thấy ${result.summary.total_missing} bản ghi thiếu, matched ${result.summary.total_matched} ảnh`, 'success');
            showToast(`Tìm thấy ${result.summary.total_missing} bản ghi thiếu`, 'success');

            document.getElementById('results-card').style.display = 'block';
        } else {
            addLog(`❌ Lỗi: ${result.error}`, 'error');
            showToast(result.error || 'Lỗi phân tích', 'error');
        }
    } catch (error) {
        stopLogStream();
        addLog(`❌ Lỗi: ${error.message}`, 'error');
        showToast('Lỗi: ' + error.message, 'error');
    } finally {
        btn.disabled = false;
        btn.textContent = '🚀 Bắt Đầu Phân Tích';
        setTimeout(() => {
            progressSection.style.display = 'none';
        }, 2000);
    }
}

function updateProgress(title, percent) {
    document.getElementById('progress-title').textContent = title;
    document.getElementById('progress-percent').textContent = `${percent}%`;
    document.getElementById('progress-fill').style.width = `${percent}%`;
}

function displayResults(records) {
    const tbody = document.getElementById('analysis-table-body');

    if (!records || records.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="empty-message">Không tìm thấy bản ghi nào thiếu dữ liệu ✅</td></tr>';
        return;
    }

    tbody.innerHTML = records.map((record, index) => `
        <tr>
            <td>${index + 1}</td>
            <td>${record.person_name}</td>
            <td>${record.date}</td>
            <td>${record.weekday}</td>
            <td>${record.issue_description}</td>
            <td>${record.matched_image
            ? `<img src="/matched-image/${encodeURIComponent(record.matched_image)}" 
                       alt="Matched" style="width:60px;height:60px;object-fit:cover;border-radius:4px;" 
                       onerror="this.style.display='none';this.nextSibling.style.display='block'">
                   <span style="display:none;color:#999;">Không có</span>`
            : '<span style="color:#999;">Không có</span>'
        }</td>
        </tr>
    `).join('');
}

// ==================== Export ====================

async function exportWord() {
    if (!analysisResults) {
        showToast('Vui lòng chạy phân tích trước', 'warning');
        return;
    }

    const projectName = document.getElementById('project-name').value.trim();
    const month = document.getElementById('export-month').value.trim();

    showToast('Đang xuất file Word...', 'info');

    try {
        const result = await apiPost('/api/export-word', {
            project_name: projectName,
            month: month,
            records: analysisResults.records
        });

        if (result.success) {
            showToast(`Đã xuất file: ${result.filename}`, 'success');
            loadResultFiles();
        } else {
            showToast(result.error || 'Lỗi xuất file', 'error');
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
    }
}

// ==================== Results ====================

async function loadResultFiles() {
    try {
        const data = await apiGet('/api/files/results');
        const tbody = document.getElementById('results-table-body');

        if (data.files && data.files.length > 0) {
            tbody.innerHTML = data.files.map((file, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td>${file.name}</td>
                    <td>${formatFileSize(file.size)}</td>
                    <td>
                        <a href="/api/files/download/${file.name}" class="btn btn-primary btn-sm">
                            📥 Tải về
                        </a>
                    </td>
                </tr>
            `).join('');
        } else {
            tbody.innerHTML = '<tr><td colspan="4" class="empty-message">Chưa có file kết quả nào</td></tr>';
        }
    } catch (error) {
        console.error('Lỗi load result files:', error);
    }
}

// ==================== PDF Extraction ====================

// PDF Upload Zone
document.addEventListener('DOMContentLoaded', () => {
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

    loadResultFiles();
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
    btn.textContent = '⏳ Đang xử lý...';
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
            btn.textContent = '🚀 Bắt Đầu Tách';
        }
    } catch (error) {
        showToast('Lỗi: ' + error.message, 'error');
        btn.disabled = false;
        btn.textContent = '🚀 Bắt Đầu Tách';
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
            document.getElementById('btn-extract').textContent = '🚀 Bắt Đầu Tách';
            loadPDFExtractedFiles();

            setTimeout(() => {
                document.getElementById('pdf-progress-section').style.display = 'none';
            }, 2000);
        } else if (result.status === 'error') {
            showToast('Lỗi: ' + result.error, 'error');
            document.getElementById('btn-extract').disabled = false;
            document.getElementById('btn-extract').textContent = '🚀 Bắt Đầu Tách';
        } else {
            // Continue checking
            setTimeout(checkPDFProgress, 1000);
        }
    } catch (error) {
        console.error('Lỗi kiểm tra tiến độ:', error);
        setTimeout(checkPDFProgress, 2000);
    }
}

async function loadPDFUploads() {
    try {
        const data = await apiGet('/api/pdf/uploads');
        const tbody = document.getElementById('pdf-uploads-table-body');

        if (data.files && data.files.length > 0) {
            document.getElementById('stat-pdf-uploads').textContent = data.files.length;
            tbody.innerHTML = data.files.map((file, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td>${file.name}</td>
                    <td>${formatFileSize(file.size)}</td>
                    <td>
                        <button class="btn btn-success btn-sm" onclick="selectPDFForExtract('${file.name}')">
                            🚀 Tách
                        </button>
                    </td>
                </tr>
            `).join('');
        } else {
            document.getElementById('stat-pdf-uploads').textContent = '0';
            tbody.innerHTML = '<tr><td colspan="4" class="empty-message">Chưa có file PDF nào</td></tr>';
        }
    } catch (error) {
        console.error('Lỗi load PDF uploads:', error);
    }
}

function selectPDFForExtract(filename) {
    pdfFilename = filename;
    document.getElementById('pdf-filename').textContent = filename;
    document.getElementById('pdf-selected-file').style.display = 'block';
    showToast(`Đã chọn: ${filename}`, 'info');
}

async function loadPDFExtractedFiles() {
    try {
        const data = await apiGet('/api/pdf/files');
        const container = document.getElementById('pdf-extracted-files');

        if (data.folders && data.folders.length > 0) {
            document.getElementById('stat-pdf-extracted').textContent = data.folders.length;

            let html = '';
            for (const folder of data.folders) {
                html += `
                    <div class="card" style="margin-bottom: 16px;">
                        <div class="card-header">
                            <h3>📁 ${folder.folder} (${folder.count} files)</h3>
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
                                        ${folder.files.map((file, idx) => `
                                            <tr>
                                                <td>${idx + 1}</td>
                                                <td>${file.name}</td>
                                                <td>${formatFileSize(file.size)}</td>
                                                <td>
                                                    <a href="/api/pdf/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}" 
                                                       class="btn btn-primary btn-sm">📥 Tải về</a>
                                                </td>
                                            </tr>
                                        `).join('')}
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
}

// ==================== Utilities ====================

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function refreshData() {
    showToast('Đang làm mới...', 'info');
    await loadResultFiles();
    showToast('Đã làm mới dữ liệu', 'success');
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
        addLog('❌ Không thể kết nối log stream', 'error');
    }
}

function stopLogStream() {
    if (logEventSource) {
        logEventSource.close();
        logEventSource = null;
    }
}

