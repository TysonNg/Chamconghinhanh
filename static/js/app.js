/**
 * JavaScript cho pháº§n má»m cháº¥m cÃ´ng v2.0
 * ÄÆ¡n giáº£n hÃ³a vá»›i 3 tab: PhÃ¢n TÃ­ch + TÃ¡ch PDF + Káº¿t Quáº£
 */

// ==================== State ====================
let analysisResults = null;
let pdfTaskId = null;
let pdfFilename = null;
let logEventSource = null;
let excelTaskId = null;
let excelFilename = null;
let excelFaceTaskId = null;

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

    const icons = { success: 'âœ…', error: 'âŒ', warning: 'âš ï¸', info: 'â„¹ï¸' };
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
            analyze: 'PhÃ¢n TÃ­ch Cháº¥m CÃ´ng',
            pdf: 'TÃ¡ch PDF ThÃ nh Word',
            excel: 'TÃ¡ch Excel Cháº¥m CÃ´ng',
            results: 'Káº¿t Quáº£'
        };
        document.querySelector('.page-title').textContent = titles[tabId] || 'PhÃ¢n TÃ­ch';

        if (tabId === 'results') loadResultFiles();
        if (tabId === 'pdf') {
            loadPDFUploads();
            loadPDFExtractedFiles();
        }
        if (tabId === 'excel') {
            loadExcelUploads();
            loadExcelExtractedFiles();
            loadExcelFaceFiles();
        }
    });
});

// ==================== Analysis ====================

async function runAnalysis() {
    const btn = document.getElementById('btn-analyze');
    const progressSection = document.getElementById('progress-section');
    const logPanel = document.getElementById('log-panel');

    btn.disabled = true;
    btn.textContent = 'â³ Äang xá»­ lÃ½...';
    progressSection.style.display = 'block';
    logPanel.style.display = 'block';
    clearLogs();

    updateProgress('Äang quÃ©t file cháº¥m cÃ´ng...', 10);
    addLog('ðŸš€ Báº¯t Ä‘áº§u phÃ¢n tÃ­ch...', 'info');

    // Connect to SSE for real-time logs
    startLogStream();

    try {
        // Step 1: PhÃ¢n tÃ­ch cháº¥m cÃ´ng
        updateProgress('Äang phÃ¢n tÃ­ch ngÃ y thiáº¿u...', 30);
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
            updateProgress('Äang hiá»ƒn thá»‹ káº¿t quáº£...', 90);
            displayResults(result.records);

            updateProgress('HoÃ n thÃ nh!', 100);
            addLog(`âœ… HoÃ n thÃ nh! TÃ¬m tháº¥y ${result.summary.total_missing} báº£n ghi thiáº¿u, matched ${result.summary.total_matched} áº£nh`, 'success');
            showToast(`TÃ¬m tháº¥y ${result.summary.total_missing} báº£n ghi thiáº¿u`, 'success');

            document.getElementById('results-card').style.display = 'block';
        } else {
            addLog(`âŒ Lá»—i: ${result.error}`, 'error');
            showToast(result.error || 'Lá»—i phÃ¢n tÃ­ch', 'error');
        }
    } catch (error) {
        stopLogStream();
        addLog(`âŒ Lá»—i: ${error.message}`, 'error');
        showToast('Lá»—i: ' + error.message, 'error');
    } finally {
        btn.disabled = false;
        btn.textContent = 'ðŸš€ Báº¯t Äáº§u PhÃ¢n TÃ­ch';
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
        tbody.innerHTML = '<tr><td colspan="6" class="empty-message">KhÃ´ng tÃ¬m tháº¥y báº£n ghi nÃ o thiáº¿u dá»¯ liá»‡u âœ…</td></tr>';
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
                   <span style="display:none;color:#999;">KhÃ´ng cÃ³</span>`
            : '<span style="color:#999;">KhÃ´ng cÃ³</span>'
        }</td>
        </tr>
    `).join('');
}

// ==================== Export ====================

async function exportWord() {
    if (!analysisResults) {
        showToast('Vui lÃ²ng cháº¡y phÃ¢n tÃ­ch trÆ°á»›c', 'warning');
        return;
    }

    const projectName = document.getElementById('project-name').value.trim();
    const month = document.getElementById('export-month').value.trim();

    showToast('Äang xuáº¥t file Word...', 'info');

    try {
        const result = await apiPost('/api/export-word', {
            project_name: projectName,
            month: month,
            records: analysisResults.records
        });

        if (result.success) {
            showToast(`ÄÃ£ xuáº¥t file: ${result.filename}`, 'success');
            loadResultFiles();
        } else {
            showToast(result.error || 'Lá»—i xuáº¥t file', 'error');
        }
    } catch (error) {
        showToast('Lá»—i: ' + error.message, 'error');
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
                            ðŸ“¥ Táº£i vá»
                        </a>
                    </td>
                </tr>
            `).join('');
        } else {
            tbody.innerHTML = '<tr><td colspan="4" class="empty-message">ChÆ°a cÃ³ file káº¿t quáº£ nÃ o</td></tr>';
        }
    } catch (error) {
        console.error('Lá»—i load result files:', error);
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
                showToast('Vui lÃ²ng chá»n file PDF', 'warning');
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
    showToast('Äang upload file PDF...', 'info');

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
            showToast(`ÄÃ£ upload: ${result.filename}`, 'success');
            loadPDFUploads();
        } else {
            showToast(result.error || 'Lá»—i upload', 'error');
        }
    } catch (error) {
        showToast('Lá»—i: ' + error.message, 'error');
    }
}

async function extractPDF() {
    if (!pdfFilename) {
        showToast('Vui lÃ²ng chá»n file PDF trÆ°á»›c', 'warning');
        return;
    }

    const btn = document.getElementById('btn-extract');
    const progressSection = document.getElementById('pdf-progress-section');

    btn.disabled = true;
    btn.textContent = 'â³ Äang xá»­ lÃ½...';
    progressSection.style.display = 'block';

    try {
        const result = await apiPost('/api/pdf/extract', { filename: pdfFilename });

        if (result.success) {
            pdfTaskId = result.task_id;
            showToast('ÄÃ£ báº¯t Ä‘áº§u tÃ¡ch PDF...', 'info');
            checkPDFProgress();
        } else {
            showToast(result.error || 'Lá»—i tÃ¡ch PDF', 'error');
            btn.disabled = false;
            btn.textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
        }
    } catch (error) {
        showToast('Lá»—i: ' + error.message, 'error');
        btn.disabled = false;
        btn.textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
    }
}

async function checkPDFProgress() {
    if (!pdfTaskId) return;

    try {
        const result = await apiGet(`/api/pdf/status/${pdfTaskId}`);

        // Update progress
        document.getElementById('pdf-progress-title').textContent = result.message || 'Äang xá»­ lÃ½...';
        document.getElementById('pdf-progress-percent').textContent = `${result.progress}%`;
        document.getElementById('pdf-progress-fill').style.width = `${result.progress}%`;
        document.getElementById('pdf-progress-detail').textContent =
            result.current_page ? `Trang ${result.current_page}/${result.total}` : '';

        if (result.status === 'completed') {
            showToast(`HoÃ n thÃ nh! ÄÃ£ táº¡o ${result.files_created.length} file Word.`, 'success');
            document.getElementById('btn-extract').disabled = false;
            document.getElementById('btn-extract').textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
            loadPDFExtractedFiles();

            setTimeout(() => {
                document.getElementById('pdf-progress-section').style.display = 'none';
            }, 2000);
        } else if (result.status === 'error') {
            showToast('Lá»—i: ' + result.error, 'error');
            document.getElementById('btn-extract').disabled = false;
            document.getElementById('btn-extract').textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
        } else {
            // Continue checking
            setTimeout(checkPDFProgress, 1000);
        }
    } catch (error) {
        console.error('Lá»—i kiá»ƒm tra tiáº¿n Ä‘á»™:', error);
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
                            ðŸš€ TÃ¡ch
                        </button>
                    </td>
                </tr>
            `).join('');
        } else {
            document.getElementById('stat-pdf-uploads').textContent = '0';
            tbody.innerHTML = '<tr><td colspan="4" class="empty-message">ChÆ°a cÃ³ file PDF nÃ o</td></tr>';
        }
    } catch (error) {
        console.error('Lá»—i load PDF uploads:', error);
    }
}

function selectPDFForExtract(filename) {
    pdfFilename = filename;
    document.getElementById('pdf-filename').textContent = filename;
    document.getElementById('pdf-selected-file').style.display = 'block';
    showToast(`ÄÃ£ chá»n: ${filename}`, 'info');
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
                            <h3>ðŸ“ ${folder.folder} (${folder.count} files)</h3>
                        </div>
                        <div class="card-body">
                            <div class="table-container">
                                <table class="data-table">
                                    <thead>
                                        <tr>
                                            <th>STT</th>
                                            <th>TÃªn File</th>
                                            <th>KÃ­ch ThÆ°á»›c</th>
                                            <th>Thao TÃ¡c</th>
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
                                                       class="btn btn-primary btn-sm">ðŸ“¥ Táº£i vá»</a>
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
            container.innerHTML = '<p class="empty-message">ChÆ°a cÃ³ file nÃ o Ä‘Æ°á»£c tÃ¡ch</p>';
        }
    } catch (error) {
        console.error('Lá»—i load extracted files:', error);
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
    showToast('Äang lÃ m má»›i...', 'info');
    await loadResultFiles();
    showToast('ÄÃ£ lÃ m má»›i dá»¯ liá»‡u', 'success');
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
            addLog('ðŸ”Œ Káº¿t ná»‘i log stream...', 'info');
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
        addLog('âŒ KhÃ´ng thá»ƒ káº¿t ná»‘i log stream', 'error');
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
        showToast('Vui lÃ²ng chá»n file .xls hoáº·c .xlsx', 'warning');
        return;
    }

    showToast('Äang upload file Excel...', 'info');
    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/excel/upload', { method: 'POST', body: formData });
        const result = await response.json();
        if (result.success) {
            excelFilename = result.filename;
            document.getElementById('excel-filename').textContent = result.filename;
            document.getElementById('excel-selected-file').style.display = 'block';
            showToast(`ÄÃ£ upload: ${result.filename}`, 'success');
            loadExcelUploads();
        } else {
            showToast(result.error || 'Lá»—i upload', 'error');
        }
    } catch (error) {
        showToast('Lá»—i: ' + error.message, 'error');
    }
}

async function extractExcel() {
    if (!excelFilename) {
        showToast('Vui lÃ²ng chá»n file Excel trÆ°á»›c', 'warning');
        return;
    }
    const btn = document.getElementById('btn-excel-extract');
    const progressSection = document.getElementById('excel-progress-section');
    btn.disabled = true;
    btn.textContent = 'â³ Äang xá»­ lÃ½...';
    progressSection.style.display = 'block';
    document.getElementById('excel-progress-fill').style.width = '0%';
    document.getElementById('excel-progress-percent').textContent = '0%';
    document.getElementById('excel-progress-title').textContent = 'Äang khá»Ÿi Ä‘á»™ng...';

    try {
        const result = await apiPost('/api/excel/extract', { filename: excelFilename });
        if (result.success) {
            excelTaskId = result.task_id;
            showToast('ÄÃ£ báº¯t Ä‘áº§u tÃ¡ch Excel...', 'info');
            checkExcelProgress();
        } else {
            showToast(result.error || 'Lá»—i tÃ¡ch Excel', 'error');
            btn.disabled = false;
            btn.textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
        }
    } catch (error) {
        showToast('Lá»—i: ' + error.message, 'error');
        btn.disabled = false;
        btn.textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
    }
}

async function checkExcelProgress() {
    if (!excelTaskId) return;
    try {
        const result = await apiGet(`/api/excel/status/${excelTaskId}`);
        const pct = result.total > 0 ? Math.round((result.progress / result.total) * 100) : 0;
        document.getElementById('excel-progress-title').textContent =
            result.current ? `Äang xuáº¥t: ${result.current}` : 'Äang xá»­ lÃ½...';
        document.getElementById('excel-progress-percent').textContent = `${pct}%`;
        document.getElementById('excel-progress-fill').style.width = `${pct}%`;
        document.getElementById('excel-progress-detail').textContent =
            `${result.progress}/${result.total} ngÆ°á»i`;

        if (result.status === 'completed') {
            document.getElementById('stat-excel-persons').textContent = result.files.length;
            document.getElementById('stat-excel-absent').textContent = 0;
            showToast(`HoÃ n thÃ nh! ÄÃ£ táº¡o ${result.files.length} file Word in theo tá»«ng ngÆ°á»i.`, 'success');
            document.getElementById('btn-excel-extract').disabled = false;
            document.getElementById('btn-excel-extract').textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
            loadExcelExtractedFiles();
            setTimeout(() => {
                document.getElementById('excel-progress-section').style.display = 'none';
            }, 3000);
        } else if (result.status === 'failed') {
            showToast('Lá»—i: ' + (result.errors[0] || 'KhÃ´ng rÃµ'), 'error');
            document.getElementById('btn-excel-extract').disabled = false;
            document.getElementById('btn-excel-extract').textContent = 'ðŸš€ Báº¯t Äáº§u TÃ¡ch';
        } else {
            setTimeout(checkExcelProgress, 800);
        }
    } catch (error) {
        console.error('Lá»—i kiá»ƒm tra tiáº¿n Ä‘á»™ Excel:', error);
        setTimeout(checkExcelProgress, 2000);
    }
}

async function loadExcelUploads() {
    try {
        const data = await apiGet('/api/excel/uploads');
        const tbody = document.getElementById('excel-uploads-table-body');
        const count = data.files ? data.files.length : 0;
        document.getElementById('stat-excel-uploads').textContent = count;

        if (data.files && data.files.length > 0) {
            tbody.innerHTML = data.files.map((file, index) => `
                <tr>
                    <td>${index + 1}</td>
                    <td>${file.name}</td>
                    <td>${formatFileSize(file.size)}</td>
                    <td>
                        <button class="btn btn-success btn-sm" onclick="selectExcelForExtract('${file.name.replace(/'/g, "\\'")}')">ðŸš€ TÃ¡ch</button>
                    </td>
                </tr>
            `).join('');
        } else {
            tbody.innerHTML = '<tr><td colspan="4" class="empty-message">ChÆ°a cÃ³ file Excel nÃ o</td></tr>';
        }
    } catch (error) {
        console.error('Lá»—i load Excel uploads:', error);
    }
}

function selectExcelForExtract(filename) {
    excelFilename = filename;
    document.getElementById('excel-filename').textContent = filename;
    document.getElementById('excel-selected-file').style.display = 'block';
    showToast(`ÄÃ£ chá»n: ${filename}`, 'info');
}

async function loadExcelExtractedFiles() {
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
                                <h3 style="margin:0;">📁 ${folder.folder} (${folder.count} người)</h3>
                                <button class="btn btn-primary btn-sm" onclick="startExcelFaceAnalyze('${safeFolder}')">🔍 Quét mặt</button>
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
                                        ${folder.files.map((file, idx) => `
                                            <tr>
                                                <td>${idx + 1}</td>
                                                <td>📄 ${file.name}</td>
                                                <td>${formatFileSize(file.size)}</td>
                                                <td>
                                                    <a href="/api/excel/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}"
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
            document.getElementById('stat-excel-persons').textContent = totalFiles;
            container.innerHTML = html;
        } else {
            container.innerHTML = '<p class="empty-message">Chưa có file Word nào được tạo</p>';
        }
    } catch (error) {
        console.error('Lỗi load Excel extracted files:', error);
    }
}

async function startExcelFaceAnalyze(folder) {
    if (!folder) {
        showToast('Vui lòng chọn thư mục đã tách', 'warning');
        return;
    }

    const thresholdInput = document.getElementById('excel-face-threshold');
    const threshold = thresholdInput ? thresholdInput.value.trim() : '';

    const progressSection = document.getElementById('excel-face-progress-section');
    progressSection.style.display = 'block';
    document.getElementById('excel-face-progress-fill').style.width = '0%';
    document.getElementById('excel-face-progress-percent').textContent = '0%';
    document.getElementById('excel-face-progress-title').textContent = 'Đang khởi động quét mặt...';
    document.getElementById('excel-face-progress-detail').textContent = '';

    try {
        const result = await apiPost('/api/excel/face/analyze', {
            folder,
            distance_threshold: threshold
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
            showToast(`Hoàn thành! Đã tạo ${result.files.length} file Word.`, 'success');
            loadExcelFaceFiles();
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

async function loadExcelFaceFiles() {
    try {
        const data = await apiGet('/api/excel/face/files');
        const container = document.getElementById('excel-face-files');
        if (data.folders && data.folders.length > 0) {
            let html = '';
            for (const folder of data.folders) {
                html += `
                    <div class="card" style="margin-bottom: 16px;">
                        <div class="card-header">
                            <h3>📁 ${folder.folder} (${folder.count} file)</h3>
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
                                        ${folder.files.map((file, idx) => `
                                            <tr>
                                                <td>${idx + 1}</td>
                                                <td>📄 ${file.name}</td>
                                                <td>${formatFileSize(file.size)}</td>
                                                <td>
                                                    <a href="/api/excel/face/download/${encodeURIComponent(folder.folder)}/${encodeURIComponent(file.name)}"
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
            container.innerHTML = '<p class="empty-message">Chưa có file Word nào</p>';
        }
    } catch (error) {
        console.error('Lỗi load Excel face files:', error);
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

