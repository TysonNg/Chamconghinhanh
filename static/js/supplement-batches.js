/* ==================== FAKE PHOTO SUPPLEMENT ENGINE ==================== */
(() => {
  'use strict';

  // State
  let projectsList = [];
  let currentProject = '';
  let employeesList = [];
  let currentSelectedEmployee = null;
  let selectedPhotos = []; // [{ file, target_date, target_time, thumbUrl }]
  let stagingPhotos = [];

  // Helper DOM & Notifications
  const el = (tag, text, className) => {
    const n = document.createElement(tag);
    if (text !== undefined && text !== null) n.textContent = text;
    if (className) n.className = className;
    return n;
  };

  const notify = (text, type = 'info') => {
    const msgEl = document.getElementById('batch-message');
    if (msgEl) {
      msgEl.textContent = text;
      msgEl.className = 'supplement-status-msg ' + (type === 'error' ? 'error' : type === 'success' ? 'success' : '');
    }
    const toastContainer = document.getElementById('toast-container');
    if (toastContainer) {
      const toast = el('div', text, `toast-item ${type}`);
      toastContainer.appendChild(toast);
      setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.3s ease';
        setTimeout(() => toast.remove(), 300);
      }, 3500);
    }
  };

  const getTodayISO = () => {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };

  function formatDateDisplay(isoDate) {
    if (!isoDate) return '';
    const parts = isoDate.split('-');
    if (parts.length === 3) return `${parts[2]}/${parts[1]}/${parts[0]}`;
    return isoDate;
  }

  function getRandomShiftTime(shift = 'morning') {
    let h, m, s;
    if (shift === 'afternoon' || shift === 'chieu') {
      // Ca Chiều: 17:45 - 18:15
      if (Math.random() < 0.5) {
        h = '17';
        m = String(45 + Math.floor(Math.random() * 15)).padStart(2, '0');
      } else {
        h = '18';
        m = String(Math.floor(Math.random() * 16)).padStart(2, '0');
      }
    } else {
      // Ca Sáng: 05:45 - 06:15
      if (Math.random() < 0.5) {
        h = '05';
        m = String(45 + Math.floor(Math.random() * 15)).padStart(2, '0');
      } else {
        h = '06';
        m = String(Math.floor(Math.random() * 16)).padStart(2, '0');
      }
    }
    s = String(Math.floor(Math.random() * 60)).padStart(2, '0');
    return `${h}:${m}:${s}`;
  }

  // --- 1. PROJECTS & EMPLOYEES LOADING ---
  async function loadProjects() {
    const projSelect = document.getElementById('supp-project-select');
    if (!projSelect) return;

    try {
      const res = await fetch('/api/projects').then(r => r.json());
      if (res.success && Array.isArray(res.projects)) {
        projectsList = res.projects;
        projSelect.replaceChildren();

        const defaultName = window.currentProjectName || res.default_project || (projectsList[0] ? projectsList[0].name : '');
        currentProject = defaultName;

        projectsList.forEach(p => {
          const opt = document.createElement('option');
          opt.value = p.name;
          opt.textContent = `${p.name} (${p.employee_count || 0} NV)`;
          if (p.name === currentProject) opt.selected = true;
          projSelect.appendChild(opt);
        });

        if (projSelect.value) {
          currentProject = projSelect.value;
          await loadEmployees(currentProject);
        }
      }
    } catch (e) {
      console.error('Lỗi tải danh sách dự án:', e);
    }
  }

  async function loadEmployees(projectName) {
    const hiddenVal = document.getElementById('supp-employee-val');
    const triggerLabel = document.getElementById('supp-emp-combobox-label');
    const portraitCard = document.getElementById('supp-emp-portrait-card');

    if (triggerLabel) triggerLabel.textContent = '-- Đang tải nhân viên... --';
    if (hiddenVal) hiddenVal.value = '';
    if (portraitCard) portraitCard.style.display = 'none';
    currentSelectedEmployee = null;

    try {
      const res = await fetch(`/api/portraits?project=${encodeURIComponent(projectName)}`).then(r => r.json());
      if (res.success && Array.isArray(res.employees)) {
        employeesList = res.employees;

        if (employeesList.length > 0) {
          // Select first employee by default
          selectEmployee(employeesList[0]);
        } else {
          if (triggerLabel) triggerLabel.textContent = '-- Không có nhân viên trong dự án --';
        }
        renderComboboxOptions(employeesList);
      } else {
        employeesList = [];
        if (triggerLabel) triggerLabel.textContent = '-- Lỗi tải nhân viên --';
        renderComboboxOptions([]);
      }
    } catch (e) {
      console.error('Lỗi tải nhân viên:', e);
      if (triggerLabel) triggerLabel.textContent = '-- Lỗi nạp nhân viên --';
    }
  }

  function selectEmployee(emp) {
    if (!emp) return;
    currentSelectedEmployee = emp;

    const hiddenVal = document.getElementById('supp-employee-val');
    const triggerLabel = document.getElementById('supp-emp-combobox-label');
    if (hiddenVal) hiddenVal.value = emp.name;
    if (triggerLabel) triggerLabel.textContent = `${emp.name} (${emp.image_count || 0} ảnh mẫu)`;

    // Update 90x90px Large Portrait Card
    const card = document.getElementById('supp-emp-portrait-card');
    const largeImg = document.getElementById('supp-emp-large-img');
    const nameEl = document.getElementById('supp-emp-card-name');
    const infoEl = document.getElementById('supp-emp-card-info');
    const countInput = document.getElementById('supp-random-count');

    if (card && largeImg && nameEl && infoEl) {
      nameEl.textContent = emp.name;
      infoEl.textContent = `Kho ảnh mẫu: ${emp.image_count || 0} ảnh chân dung`;
      if (countInput) {
        countInput.value = '1';
      }

      if (emp.avatar_url) {
        largeImg.src = emp.avatar_url;
        card.style.display = 'flex';
      } else {
        card.style.display = 'none';
      }
    }

    // Close menu if open
    closeComboboxMenu();
  }

  // --- 2. CUSTOM SEARCHABLE COMBOBOX LOGIC ---
  function renderComboboxOptions(listToRender) {
    const listContainer = document.getElementById('supp-emp-options-list');
    if (!listContainer) return;
    listContainer.replaceChildren();

    if (!listToRender || listToRender.length === 0) {
      const emptyDiv = el('div', 'Không tìm thấy nhân viên phù hợp', 'supp-combobox-empty');
      listContainer.appendChild(emptyDiv);
      return;
    }

    listToRender.forEach(emp => {
      const opt = el('div', null, 'supp-combobox-option');
      if (currentSelectedEmployee && currentSelectedEmployee.employee_id === emp.employee_id) {
        opt.classList.add('selected');
      }

      // Avatar
      const img = el('img', null, 'supp-combobox-opt-avatar');
      img.src = emp.avatar_url || '/static/img/default-avatar.png';
      img.alt = emp.name;

      // Name
      const nameSpan = el('span', emp.name, 'supp-combobox-opt-name');

      // Badge
      const badgeSpan = el('span', `${emp.image_count || 0} ảnh`, 'supp-combobox-opt-badge');

      opt.append(img, nameSpan, badgeSpan);

      opt.onclick = e => {
        e.stopPropagation();
        selectEmployee(emp);
      };

      listContainer.appendChild(opt);
    });
  }

  function openComboboxMenu() {
    const trigger = document.getElementById('supp-emp-combobox-trigger');
    const menu = document.getElementById('supp-emp-combobox-menu');
    const searchInp = document.getElementById('supp-emp-search-input');
    if (!trigger || !menu) return;

    trigger.classList.add('open');
    menu.style.display = 'flex';
    if (searchInp) {
      searchInp.value = '';
      renderComboboxOptions(employeesList);
      setTimeout(() => searchInp.focus(), 50);
    }
  }

  function closeComboboxMenu() {
    const trigger = document.getElementById('supp-emp-combobox-trigger');
    const menu = document.getElementById('supp-emp-combobox-menu');
    if (trigger) trigger.classList.remove('open');
    if (menu) menu.style.display = 'none';
  }

  function setupEmployeeCombobox() {
    const trigger = document.getElementById('supp-emp-combobox-trigger');
    const menu = document.getElementById('supp-emp-combobox-menu');
    const searchInp = document.getElementById('supp-emp-search-input');
    const largeThumb = document.getElementById('supp-emp-large-thumb');

    if (trigger) {
      trigger.onclick = e => {
        e.stopPropagation();
        if (trigger.classList.contains('open')) {
          closeComboboxMenu();
        } else {
          openComboboxMenu();
        }
      };
    }

    if (searchInp) {
      searchInp.onclick = e => e.stopPropagation();
      searchInp.oninput = () => {
        const q = searchInp.value.trim().toLowerCase();
        if (!q) {
          renderComboboxOptions(employeesList);
        } else {
          const filtered = employeesList.filter(emp => emp.name.toLowerCase().includes(q));
          renderComboboxOptions(filtered);
        }
      };
    }

    // Lightbox zoom on clicking large 90x90px portrait
    if (largeThumb) {
      largeThumb.onclick = () => {
        if (currentSelectedEmployee && currentSelectedEmployee.avatar_url) {
          if (typeof window.openLightbox === 'function') {
            window.openLightbox(
              currentSelectedEmployee.avatar_url,
              `${currentSelectedEmployee.name} - Ảnh chân dung mẫu (${currentProject})`
            );
          } else {
            window.open(currentSelectedEmployee.avatar_url, '_blank');
          }
        }
      };
    }

    // Click outside closes combobox
    document.addEventListener('click', e => {
      const container = document.getElementById('supp-combobox-container');
      if (container && !container.contains(e.target)) {
        closeComboboxMenu();
      }
    });
  }

  // --- 3. RANDOM PORTRAIT PICKER ---
  function setupRandomPortraitPicker() {
    const btnPick = document.getElementById('btn-pick-random-portrait');
    const countInput = document.getElementById('supp-random-count');
    const appendCheckbox = document.getElementById('chk-append-portrait');

    if (!btnPick) return;

    btnPick.onclick = async () => {
      if (!currentSelectedEmployee) {
        notify('Vui lòng chọn nhân viên trước.', 'warning');
        return;
      }

      const images = currentSelectedEmployee.images || [];
      if (images.length === 0) {
        notify(`Nhân viên ${currentSelectedEmployee.name} không có ảnh chân dung mẫu nào trong hệ thống.`, 'warning');
        return;
      }

      const count = Number(countInput ? countInput.value : '1');
      if (!Number.isInteger(count) || count < 1 || count > images.length) {
        notify(`Nhập số nguyên từ 1 đến ${images.length}. Kho hiện có ${images.length} ảnh có thể lấy.`, 'warning');
        return;
      }
      const isAppend = appendCheckbox ? appendCheckbox.checked : false;

      btnPick.disabled = true;
      const originalText = btnPick.innerHTML;
      btnPick.innerHTML = `<span>Đang bốc ${count} ảnh...</span>`;

      try {
        // Randomly pick `count` filenames from images
        const pool = [...images];
        const pickedFilenames = [];
        for (let i = 0; i < count; i++) {
          if (pool.length > 0) {
            const randIdx = Math.floor(Math.random() * pool.length);
            pickedFilenames.push(pool.splice(randIdx, 1)[0]);
          } else {
            // Repeat if count > pool size
            const randIdx = Math.floor(Math.random() * images.length);
            pickedFilenames.push(images[randIdx]);
          }
        }

        const newItems = [];
        const todayStr = getTodayISO();

        for (let i = 0; i < pickedFilenames.length; i++) {
          const fname = pickedFilenames[i];
          const url = currentSelectedEmployee.image_urls[fname];
          const blob = await fetch(url).then(r => r.blob());
          const file = new File([blob], fname.split('/').pop(), { type: blob.type || 'image/jpeg' });
          const thumbUrl = URL.createObjectURL(blob);

          // Random time in morning shift: 07:45 - 08:15
          const isLate = Math.random() > 0.5;
          const h = isLate ? '08' : '07';
          const m = isLate ? String(Math.floor(Math.random() * 16)).padStart(2, '0') : String(45 + Math.floor(Math.random() * 15)).padStart(2, '0');
          const s = String(Math.floor(Math.random() * 60)).padStart(2, '0');

          newItems.push({
            file,
            target_date: todayStr,
            target_time: `${h}:${m}:${s}`,
            thumbUrl
          });
        }

        if (isAppend) {
          selectedPhotos = [...selectedPhotos, ...newItems];
        } else {
          selectedPhotos = newItems;
        }

        renderSelectedGrid();
        notify(`✓ Đã nạp thành công ${newItems.length} ảnh chân dung của ${currentSelectedEmployee.name}!`, 'success');
      } catch (err) {
        console.error('Lỗi bốc ảnh chân dung:', err);
        notify('Lỗi nạp ảnh chân dung: ' + err.message, 'error');
      } finally {
        btnPick.disabled = false;
        btnPick.innerHTML = originalText;
      }
    };
  }

  // --- 4. MULTI-PHOTO SELECTION & GRID ---
  function handleFilesSelected(files) {
    if (!files || files.length === 0) return;

    const newFiles = Array.from(files);
    if (selectedPhotos.length + newFiles.length > 31) {
      notify('Tối đa 31 ảnh cùng lúc. Vui lòng chọn ít ảnh hơn.', 'warning');
      return;
    }

    const todayStr = getTodayISO();
    const timeStr = '08:00:00';

    newFiles.forEach(file => {
      if (file.size > 20 * 1024 * 1024) {
        notify(`File ${file.name} vượt quá 20MB và bị bỏ qua.`, 'warning');
        return;
      }
      const thumbUrl = URL.createObjectURL(file);
      selectedPhotos.push({
        file,
        target_date: todayStr,
        target_time: timeStr,
        thumbUrl
      });
    });

    renderSelectedGrid();
  }

  function renderSelectedGrid() {
    const section = document.getElementById('supp-selected-section');
    const grid = document.getElementById('supp-selected-grid');
    const countSpan = document.getElementById('supp-selected-count');
    const sameDateInput = document.getElementById('tool-same-date-input');

    if (!section || !grid) return;

    if (selectedPhotos.length === 0) {
      section.style.display = 'none';
      grid.replaceChildren();
      const fileInput = document.getElementById('supp-photos-input');
      if (fileInput) fileInput.value = '';
      return;
    }

    section.style.display = 'block';
    if (countSpan) countSpan.textContent = selectedPhotos.length;
    if (sameDateInput && !sameDateInput.value) {
      sameDateInput.value = selectedPhotos[0].target_date || getTodayISO();
    }

    grid.replaceChildren();

    selectedPhotos.forEach((item, index) => {
      const card = el('div', null, 'supp-photo-item-card');

      // Remove button
      const removeBtn = el('button', '×', 'supp-item-remove-btn');
      removeBtn.type = 'button';
      removeBtn.title = 'Bỏ ảnh này';
      removeBtn.onclick = () => {
        selectedPhotos.splice(index, 1);
        renderSelectedGrid();
      };

      // Thumbnail
      const thumb = el('img', null, 'supp-photo-item-thumb');
      thumb.src = item.thumbUrl;
      thumb.alt = item.file.name;
      thumb.style.cursor = 'pointer';
      thumb.title = 'Nhấp để xem phóng to';
      thumb.onclick = () => {
        if (typeof window.openLightbox === 'function') {
          window.openLightbox(item.thumbUrl, `${item.file.name} - Ngày đề nghị bổ sung: ${formatDateDisplay(item.target_date)}. Giữ nguyên ngày trên ảnh.`);
        }
      };

      // Info
      const info = el('div', null, 'supp-photo-item-info');
      const nameP = el('div', `#${index + 1}. ${item.file.name}`, 'supp-photo-item-name');
      nameP.title = item.file.name;

      // Date and Time inputs
      const inputsRow = el('div', null, 'supp-photo-item-inputs');

      if (!item.target_time) {
        const curShift = document.getElementById('tool-shift-select')?.value || 'morning';
        item.target_time = getRandomShiftTime(curShift);
      }

      // Date input row
      const dateRow = el('div', null, 'supp-photo-input-row');
      const dateLbl = el('span', 'Ngày', 'supp-photo-input-label');
      const dateInp = el('input', null, 'form-control supp-photo-input-control');
      dateInp.type = 'date';
      dateInp.required = true;
      dateInp.setAttribute('aria-label', 'Ngày đề nghị bổ sung công');
      dateInp.value = item.target_date;
      dateInp.onchange = () => { item.target_date = dateInp.value; };
      dateRow.append(dateLbl, dateInp);

      // Time input row
      const timeRow = el('div', null, 'supp-photo-input-row');
      const timeLbl = el('span', 'Giờ', 'supp-photo-input-label');
      const timeInp = el('input', null, 'form-control supp-photo-input-control');
      timeInp.type = 'text';
      timeInp.value = item.target_time;
      timeInp.placeholder = 'HH:MM:SS';
      timeInp.setAttribute('aria-label', 'Giờ đề nghị bổ sung');
      timeInp.onchange = () => { item.target_time = timeInp.value; };
      timeRow.append(timeLbl, timeInp);

      inputsRow.append(dateRow, timeRow);
      info.append(nameP, inputsRow);

      card.append(removeBtn, thumb, info);
      grid.appendChild(card);
    });
  }

  // --- 5. BATCH TOOLS (CÙNG NGÀY, LIÊN TIẾP, RANDOM GIỜ) ---
  function setupBatchTools() {
    // 1. Gán cùng ngày
    const btnSame = document.getElementById('btn-apply-same-date');
    const inpSame = document.getElementById('tool-same-date-input');
    if (btnSame && inpSame) {
      btnSame.onclick = () => {
        const val = inpSame.value;
        if (!val) {
          notify('Vui lòng chọn ngày để gán.', 'warning');
          return;
        }
        selectedPhotos.forEach(p => { p.target_date = val; });
        renderSelectedGrid();
        notify(`Đã gán ngày ${formatDateDisplay(val)} cho toàn bộ ${selectedPhotos.length} ảnh.`, 'success');
      };
    }

    // 2. Tăng ngày liên tiếp (+1)
    const btnSeq = document.getElementById('btn-apply-sequential-dates');
    if (btnSeq) {
      btnSeq.onclick = () => {
        if (selectedPhotos.length === 0) return;
        const startVal = (inpSame && inpSame.value) ? inpSame.value : selectedPhotos[0].target_date || getTodayISO();
        let cur = new Date(startVal);
        if (isNaN(cur.getTime())) cur = new Date();

        selectedPhotos.forEach((p, idx) => {
          const d = new Date(cur);
          d.setDate(cur.getDate() + idx);
          const y = d.getFullYear();
          const m = String(d.getMonth() + 1).padStart(2, '0');
          const day = String(d.getDate()).padStart(2, '0');
          p.target_date = `${y}-${m}-${day}`;
        });

        renderSelectedGrid();
        notify(`Đã phân bổ ngày liên tiếp bắt đầu từ ${formatDateDisplay(startVal)}.`, 'success');
      };
    }

    // 3. Random giờ ca
    const btnRand = document.getElementById('btn-apply-random-time');
    const selShift = document.getElementById('tool-shift-select');
    if (btnRand && selShift) {
      btnRand.onclick = () => {
        if (selectedPhotos.length === 0) return;
        const shift = selShift.value;

        selectedPhotos.forEach(p => {
          p.target_time = getRandomShiftTime(shift);
        });

        renderSelectedGrid();
        notify(`Đã tạo giờ ngẫu nhiên trong ${shift === 'morning' ? 'Ca Sáng (05:45 - 06:15)' : 'Ca Chiều (17:45 - 18:15)'} cho toàn bộ ảnh.`, 'success');
      };
    }

    // Nút hủy tất cả
    const btnClear = document.getElementById('btn-clear-all-selected');
    if (btnClear) {
      btnClear.onclick = () => {
        selectedPhotos = [];
        renderSelectedGrid();
      };
    }
  }

  // --- 6. FORM SUBMIT (BẮT ĐẦU TẠO FAKE ẢNH) ---
  function setupFormSubmit() {
    const form = document.getElementById('batch-create');
    const submitBtn = document.getElementById('btn-submit-fake-photos');
    if (!form || !submitBtn) return;

    form.onsubmit = async e => {
      e.preventDefault();
      if (selectedPhotos.length === 0) {
        notify('Vui lòng chọn hoặc bốc ít nhất 1 ảnh để xử lý.', 'warning');
        return;
      }

      const proj = document.getElementById('supp-project-select').value;
      const emp = document.getElementById('supp-employee-val').value;
      if (!proj || !emp) {
        notify('Vui lòng chọn đầy đủ Dự án và Nhân viên.', 'warning');
        return;
      }

      const fd = new FormData();
      fd.append('project', proj);
      fd.append('employee', emp);

      const selShift = document.getElementById('tool-shift-select');
      const curShift = selShift ? selShift.value : 'morning';
      fd.append('shift', curShift);

      const configs = selectedPhotos.map(p => ({
        target_date: p.target_date,
        target_time: p.target_time || getRandomShiftTime(curShift)
      }));
      fd.append('configs', JSON.stringify(configs));

      selectedPhotos.forEach(p => {
        fd.append('photos', p.file);
      });

      submitBtn.disabled = true;
      submitBtn.innerHTML = `
        <span class="spinner-border spinner-border-sm" role="status" aria-hidden="true"></span>
        <span>Đang AI xử lý & sửa watermark ${selectedPhotos.length} ảnh...</span>
      `;

      try {
        const resp = await fetch('/api/supplement/records', {
          method: 'POST',
          body: fd
        });
        const data = await resp.json();
        if (!resp.ok || !data.success) {
          throw new Error(data.error || 'Lỗi tạo ảnh bổ sung');
        }

        notify(`Đã tạo thành công ${data.count} ảnh bổ sung với watermark và EXIF mới!`, 'success');

        // Reload staging
        await loadStaging();

        // Scroll down to staging area
        const stagingCard = document.getElementById('supp-staging-card');
        if (stagingCard) {
          stagingCard.scrollIntoView({ behavior: 'smooth' });
        }
      } catch (err) {
        notify(err.message, 'error');
      } finally {
        submitBtn.disabled = false;
        submitBtn.innerHTML = `
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
          <span>Tạo & Lưu ảnh bổ sung (AI Sửa Watermark)</span>
        `;
      }
    };
  }

  // --- 7. STAGING WORKSPACE & ACTIONS ---
  async function loadStaging() {
    const grid = document.getElementById('supp-staging-grid');
    const emptyBox = document.getElementById('supp-staging-empty');
    const countBadge = document.getElementById('staging-count');
    const actionsBar = document.getElementById('staging-actions-bar');
    if (!grid) return;

    try {
      const resp = await fetch('/api/supplement/staging').then(r => r.json());
      if (resp.success && Array.isArray(resp.items)) {
        stagingPhotos = resp.items;
        if (countBadge) countBadge.textContent = stagingPhotos.length;

        if (stagingPhotos.length === 0) {
          if (actionsBar) actionsBar.style.setProperty('display', 'none', 'important');
          if (emptyBox) emptyBox.style.display = 'flex';
          grid.replaceChildren();
          return;
        }

        if (actionsBar) actionsBar.style.setProperty('display', 'flex', 'important');
        if (emptyBox) emptyBox.style.display = 'none';

        grid.replaceChildren();

        stagingPhotos.forEach(item => {
          const card = el('div', null, 'supp-staging-card-item');

          // Thumbnail with zoom on click
          const thumbWrap = el('div', null, 'supp-staging-thumb-wrap');
          const img = el('img', null, 'supp-staging-thumb-img');
          const cacheBuster = `?v=${encodeURIComponent(item.created_at || '')}_${Date.now()}`;
          img.src = `${item.url}${cacheBuster}`;
          img.alt = item.original_name;
          img.loading = 'lazy';

          const zoomOverlay = el('div', null, 'photo-card-zoom-overlay');
          zoomOverlay.innerHTML = `
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
              <circle cx="11" cy="11" r="8"/>
              <line x1="21" y1="21" x2="16.65" y2="16.65"/>
              <line x1="11" y1="8" x2="11" y2="14"/>
              <line x1="8" y1="11" x2="14" y2="11"/>
            </svg>
            <span>Phóng to</span>
          `;

          thumbWrap.append(img, zoomOverlay);
          thumbWrap.onclick = () => {
            const cap = `${item.employee} · Ngày đề nghị: ${formatDateDisplay(item.target_date)} ${item.target_time || ''} · Watermark & EXIF: ${formatDateDisplay(item.target_date)} ${item.target_time || ''} · ${item.project}`;
            openStagingLightbox(item, cap);
          };

          // Body (Gọn gàng, súc tích, không tràn khung)
          const body = el('div', null, 'supp-staging-card-body');

          // Meta Row: Badge ngày giờ + Status tag
          const metaRow = el('div', null, 'supp-staging-meta-row');
          const badge = el('span', `📅 ${formatDateDisplay(item.target_date)} · ${item.target_time || ''}`, 'supp-staging-badge');
          const statusLabels = {
            completed: '✓ AI Khớp',
            needs_confirmation: '⚠ Chờ xác nhận',
            verification_failed: '✕ Chưa đạt hậu kiểm',
            original_missing: '✕ Thiếu ảnh gốc',
            not_requested: 'Ảnh gốc',
            legacy: 'Ảnh cũ'
          };
          const watermarkStatus = item.watermark_status || 'legacy';
          const statusTag = el(
            'span',
            statusLabels[watermarkStatus] || watermarkStatus,
            `supp-staging-status-tag ${watermarkStatus}`
          );
          metaRow.append(badge, statusTag);

          // Row: Tên nhân viên & Dự án
          const empP = el('div', null, 'supp-staging-emp-info');
          empP.title = `${item.employee} · ${item.project}`;
          const empBold = el('strong', item.employee);
          const projSpan = el('span', ` · ${item.project}`);
          empP.append(empBold, projSpan);

          // Actions Row (Bố cục 2 hàng không bao giờ bị cắt hay che chữ)
          const actions = el('div', null, 'supp-staging-card-actions');

          // Hàng 1: Nút chính Chấm Công
          const mainRow = el('div', null, 'supp-action-main-row');
          const applyBtn = el('button', 'Dùng ảnh gốc', 'btn btn-primary supp-action-main-btn');
          applyBtn.type = 'button';
          applyBtn.title = 'Chỉ dùng ảnh gốc có ngày phù hợp; giữ ảnh và lịch sử xử lý';
          applyBtn.onclick = async () => {
            applyBtn.disabled = true;
            try {
              const res = await fetch('/api/supplement/apply-to-attendance', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ids: [item.id], delete_after: true })
              }).then(r => r.json());
              if (!res.success) throw new Error(res.error || res.message || 'Ngày ảnh gốc chưa đủ điều kiện');
              if (res.success) {
                notify(`Đã đưa ảnh ngày ${formatDateDisplay(item.target_date)} vào chấm công!`, 'success');
                await loadStaging();
              }
            } catch (e) {
              notify('Lỗi: ' + e.message, 'error');
            } finally {
              applyBtn.disabled = false;
            }
          };
          mainRow.appendChild(applyBtn);

          // Hàng 2: Nhóm nút công cụ nhỏ gọn
          const toolsRow = el('div', null, 'supp-action-tools-row');

          // Tool 1: Gen lại ảnh bằng AI
          const regenBtn = el('button', '🔄 Gen lại', 'supp-tool-btn btn-regen-ai');
          regenBtn.type = 'button';
          regenBtn.title = 'Thử tạo lại ảnh này bằng AI (giữ nguyên ngày giờ & nhân viên)';
          regenBtn.onclick = () => regenerateStagingPhoto(item.id, regenBtn);

          // Tool 2: Check ảnh có phải AI không
          const checkAiBtn = el('button', '🛡️ Check AI', 'supp-tool-btn btn-check-ai');
          checkAiBtn.type = 'button';
          checkAiBtn.title = 'Kiểm tra xem ảnh có dấu vết C2PA hoặc SynthID không';
          checkAiBtn.onclick = () => checkAIPhoto(item.id, item.original_name);

          // Tool 3: Tải về
          const dlBtn = el('a', '⬇️', 'supp-tool-btn btn btn-secondary');
          dlBtn.href = `/api/supplement/download-staging/${item.id}`;
          dlBtn.title = 'Tải file ảnh về máy tính';

          // Tool 4: Xóa
          const delBtn = el('button', '🗑️', 'supp-tool-btn btn btn-danger');
          delBtn.type = 'button';
          delBtn.title = 'Xóa ảnh này khỏi danh sách';
          delBtn.onclick = async () => {
            try {
              await fetch('/api/supplement/delete-staging', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ ids: [item.id] })
              });
              notify('Đã xóa ảnh.', 'info');
              await loadStaging();
            } catch (e) {
              notify('Lỗi xóa: ' + e.message, 'error');
            }
          };

          toolsRow.append(regenBtn, checkAiBtn, dlBtn, delBtn);
          actions.append(mainRow, toolsRow);

          body.append(metaRow, empP, actions);
          card.append(thumbWrap, body);
          grid.appendChild(card);
        });
      }
    } catch (e) {
      console.error('Lỗi tải staging:', e);
    }
  }

  function setupStagingActions() {
    // 1. Đưa tất cả vào chấm công
    const btnApplyAll = document.getElementById('btn-apply-all-attendance');
    if (btnApplyAll) {
      btnApplyAll.onclick = async () => {
        if (stagingPhotos.length === 0) return;
        btnApplyAll.disabled = true;
        try {
          const res = await fetch('/api/supplement/apply-to-attendance', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: stagingPhotos.map(item => item.id), delete_after: true })
          }).then(r => r.json());
          if (!res.success) throw new Error(res.error || res.message || 'Ngày ảnh gốc chưa đủ điều kiện');
          if (res.success) {
            notify(`✓ ${res.message}`, 'success');
            await loadStaging();
          }
        } catch (e) {
          notify('Lỗi: ' + e.message, 'error');
        } finally {
          btnApplyAll.disabled = false;
        }
      };
    }

    // 2. Tải ZIP
    const btnDlZip = document.getElementById('btn-download-all-zip');
    if (btnDlZip) {
      btnDlZip.onclick = async () => {
        if (stagingPhotos.length === 0) return;
        try {
          const resp = await fetch('/api/supplement/download-staging-zip', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ ids: stagingPhotos.map(p => p.id) })
          });
          if (!resp.ok) throw new Error('Lỗi tải zip');
          const blob = await resp.blob();
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = `anh-bo-sung-${getTodayISO()}.zip`;
          document.body.appendChild(a);
          a.click();
          a.remove();
        } catch (e) {
          notify('Lỗi tải zip: ' + e.message, 'error');
        }
      };
    }

    // 3. Xóa tất cả
    const btnDelAll = document.getElementById('btn-delete-all-staging');
    if (btnDelAll) {
      btnDelAll.onclick = async () => {
        if (stagingPhotos.length === 0) return;
        if (!confirm(`Bạn có chắc chắn muốn xóa toàn bộ ${stagingPhotos.length} ảnh đang chờ?`)) return;
        try {
          await fetch('/api/supplement/delete-staging', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({})
          });
          notify('Đã xóa toàn bộ ảnh chờ.', 'info');
          await loadStaging();
        } catch (e) {
          notify('Lỗi xóa: ' + e.message, 'error');
        }
      };
    }

    // Refresh button
    const btnRefresh = document.getElementById('batch-refresh');
    if (btnRefresh) {
      btnRefresh.onclick = async () => {
        await loadProjects();
        await loadStaging();
        notify('Đã làm mới dữ liệu.', 'info');
      };
    }
  }

  // --- 7.1 STAGING LIGHTBOX & REGENERATE & CHECK AI ---
  function openStagingLightbox(item, cap) {
    const modal = document.getElementById('modal-lightbox');
    const img = document.getElementById('lightbox-img');
    const capEl = document.getElementById('lightbox-caption');
    const toolbar = document.getElementById('lightbox-toolbar');
    const btnRegen = document.getElementById('lightbox-btn-regen');
    const btnCheck = document.getElementById('lightbox-btn-check');
    const btnDl = document.getElementById('lightbox-btn-download');

    const fullUrl = `${item.url}?v=${Date.now()}`;
    if (img) img.src = fullUrl;
    if (capEl) capEl.textContent = cap;

    if (toolbar) {
      toolbar.style.display = 'flex';

      if (btnRegen) {
        btnRegen.onclick = (e) => {
          e.stopPropagation();
          regenerateStagingPhoto(item.id, btnRegen);
        };
      }
      if (btnCheck) {
        btnCheck.onclick = (e) => {
          e.stopPropagation();
          checkAIPhoto(item.id, item.original_name);
        };
      }
      if (btnDl) {
        btnDl.href = `/api/supplement/download-staging/${item.id}`;
      }
    }

    if (modal) modal.style.display = 'flex';
  }

  function showWatermarkConfirmation(photoId, ocr, triggerBtn) {
    const modal = document.getElementById('modal-watermark-confirm');
    const crop = document.getElementById('watermark-confirm-crop');
    const linesBox = document.getElementById('watermark-confirm-lines');
    const providersBox = document.getElementById('watermark-provider-results');
    const submit = document.getElementById('watermark-confirm-submit');
    if (!modal || !linesBox || !submit) return;

    const suggested = Array.isArray(ocr && ocr.suggested_lines) ? [...ocr.suggested_lines] : [];
    if (ocr && ocr.new_timestamp && suggested.length) suggested[0] = ocr.new_timestamp;
    if (crop) crop.src = `/api/supplement/staging/${photoId}/watermark-crop?v=${Date.now()}`;

    linesBox.replaceChildren();
    suggested.forEach((line, index) => {
      const row = el('div', null, 'watermark-confirm-line');
      const label = el('label', index === 0 ? 'Dòng ngày giờ mới' : `Dòng địa chỉ ${index}`, null);
      const input = el('input', null, 'form-control watermark-confirm-input');
      input.type = 'text';
      input.value = line || '';
      input.dataset.lineIndex = String(index);
      input.autocomplete = 'off';
      row.append(label, input);
      linesBox.appendChild(row);
    });

    if (providersBox) {
      const providerResults = (ocr && ocr.provider_results) || {};
      providersBox.textContent = Object.entries(providerResults).map(([provider, result]) => {
        const timestamp = result.timestamp_line || {};
        const refs = result.address_lines || result.reference_lines || [];
        const values = [timestamp.text || timestamp.original_text || '', ...refs.map(item => item.text || '')];
        return `${provider.toUpperCase()}:\n${values.filter(Boolean).join('\n')}`;
      }).join('\n\n') || 'Chỉ có một kết quả OCR; cần người dùng xác nhận.';
    }

    submit.onclick = async () => {
      const confirmedLines = Array.from(linesBox.querySelectorAll('.watermark-confirm-input'))
        .map(input => input.value.normalize('NFC').replace(/\s+/g, ' ').trim())
        .filter(Boolean);
      if (confirmedLines.length < 2) {
        notify('Cần xác nhận dòng ngày giờ và ít nhất một dòng địa chỉ.', 'error');
        return;
      }
      submit.disabled = true;
      try {
        const completed = await regenerateStagingPhoto(photoId, triggerBtn || submit, confirmedLines);
        if (completed) closeModal('modal-watermark-confirm');
      } finally {
        submit.disabled = false;
      }
    };
    modal.style.display = 'flex';
  }

  async function regenerateStagingPhoto(photoId, btn, confirmedLines = null) {
    if (!photoId) return;
    const origHtml = btn ? btn.innerHTML : '';
    if (btn) {
      btn.disabled = true;
      btn.innerHTML = '<span>⏳ Đang gen...</span>';
    }
    notify('Đang gửi yêu cầu AI render lại ảnh...', 'info');

    try {
      const res = await fetch(`/api/supplement/staging/${photoId}/regenerate`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(confirmedLines ? { confirmed_lines: confirmedLines } : {})
      });
      const data = await res.json();

      if (res.status === 422 && data.status === 'verification_failed') {
        notify('AI chưa tạo được watermark khớp tuyệt đối. Ảnh hiện tại được giữ nguyên.', 'error');
        return false;
      }
      if (res.status === 404 && data.status === 'original_missing') {
        notify('Không tìm thấy ảnh gốc bất biến; hệ thống từ chối gen chồng trên ảnh AI.', 'error');
        return false;
      }
      if (res.ok && data.success) {
        notify('Đã tạo lại ảnh bằng AI thành công!', 'success');
        await loadStaging();

        // Cập nhật ngay trong Lightbox nếu đang mở
        const lbImg = document.getElementById('lightbox-img');
        const modal = document.getElementById('modal-lightbox');
        if (modal && modal.style.display !== 'none' && lbImg) {
          lbImg.src = `/api/supplement/staging/${photoId}/image?v=${Date.now()}`;
        }
        return true;
      } else {
        notify('Lỗi khi tạo lại ảnh: ' + (data.error || data.status || 'Thất bại'), 'error');
        return false;
      }
    } catch (e) {
      notify('Lỗi kết nối khi tạo lại ảnh: ' + e.message, 'error');
      return false;
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = origHtml;
      }
    }
  }

  async function checkAIPhoto(photoId, fileName) {
    if (!photoId) return;
    notify('Đang quét kiểm tra dấu vết AI...', 'info');

    try {
      const res = await fetch('/api/supplement/check-ai', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ photo_id: photoId })
      }).then(r => r.json());

      if (res.success && res.report) {
        showAIInspectorModal(res.report, fileName);
      } else {
        notify('Không thể kiểm tra ảnh: ' + (res.error || 'Lỗi không xác định'), 'error');
      }
    } catch (e) {
      notify('Lỗi kết nối kiểm tra AI: ' + e.message, 'error');
    }
  }

  function showAIInspectorModal(report, fileName) {
    const modal = document.getElementById('modal-ai-inspector');
    if (!modal) return;

    const fnEl = document.getElementById('ai-inspector-filename');
    if (fnEl) {
      const name = fileName || report.filename || 'Ảnh kiểm tra';
      fnEl.textContent = name;
      fnEl.title = name;
    }

    const banner = document.getElementById('ai-inspector-banner');
    const icon = document.getElementById('ai-inspector-icon');
    const statusText = document.getElementById('ai-inspector-status-text');
    const desc = document.getElementById('ai-inspector-verdict-desc');
    const score = document.getElementById('ai-inspector-score');
    const scoreSub = document.getElementById('ai-inspector-score-sub');

    const isDetected = report.is_ai_detected || report.status === 'ai_detected';
    if (isDetected) {
      if (banner) banner.className = 'ai-inspector-verdict-card fail';
      if (icon) icon.textContent = '⚠️';
      if (statusText) statusText.textContent = 'PHÁT HIỆN DẤU HIỆU AI TRONG ẢNH';
      if (score) score.textContent = `${report.ai_score || 85}%`;
      if (scoreSub) {
        scoreSub.className = 'ai-score-sub fail';
        scoreSub.textContent = 'CẢNH BÁO';
      }
    } else {
      if (banner) banner.className = 'ai-inspector-verdict-card pass';
      if (icon) icon.textContent = '🛡️';
      if (statusText) statusText.textContent = 'ẢNH SẠCH - KHÔNG PHÁT HIỆN DẤU VẾT AI';
      if (score) score.textContent = '0%';
      if (scoreSub) {
        scoreSub.className = 'ai-score-sub pass';
        scoreSub.textContent = 'AN TOÀN';
      }
    }

    if (desc) desc.textContent = report.verdict || 'Đã làm sạch metadata C2PA, phá vỡ watermark ẩn SynthID và giữ EXIF camera gốc.';

    // Render list checks
    const listEl = document.getElementById('ai-inspector-checks-list');
    if (listEl) {
      listEl.replaceChildren();
      const iconMap = {
        c2pa_metadata: '📑',
        visual_watermark: '👁️',
        camera_exif: '📷',
        synthid_invisible: '🌊'
      };

      (report.checks || []).forEach(chk => {
        const item = document.createElement('div');
        item.className = 'ai-check-row';
        const chkIcon = iconMap[chk.id] || '🛡️';
        const pillText = chk.status === 'pass' ? '✓ ĐẠT' : (chk.status === 'warn' ? '⚠️ LƯU Ý' : '✕ CẢNH BÁO');

        item.innerHTML = `
          <div class="ai-check-row-left">
            <div class="ai-check-icon-box">${chkIcon}</div>
            <div>
              <div class="ai-check-name">${chk.name}</div>
              <div class="ai-check-detail">${chk.detail}</div>
            </div>
          </div>
          <span class="ai-check-pill ${chk.status}">${pillText}</span>
        `;
        listEl.appendChild(item);
      });
    }

    modal.style.display = 'flex';
  }

  // --- 8. DRAG AND DROP SETUP ---
  function setupDragAndDrop() {
    const dropzone = document.getElementById('supplement-dropzone');
    const fileInput = document.getElementById('supp-photos-input');
    if (!dropzone || !fileInput) return;

    fileInput.onchange = () => {
      handleFilesSelected(fileInput.files);
    };

    ['dragenter', 'dragover'].forEach(name => {
      dropzone.addEventListener(name, e => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropzone.addEventListener(name, e => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', e => {
      const dt = e.dataTransfer;
      if (dt && dt.files && dt.files.length) {
        handleFilesSelected(dt.files);
      }
    });

    // Project select change
    const projSelect = document.getElementById('supp-project-select');
    if (projSelect) {
      projSelect.onchange = async () => {
        currentProject = projSelect.value;
        await loadEmployees(currentProject);
      };
    }
  }

  // --- 9. AI SETTINGS & PROMPT MANAGEMENT ---
  let aiConfig = null;

  async function loadAIConfig() {
    try {
      const res = await fetch('/api/supplement/ai-config').then(r => r.json());
      aiConfig = res;
      updateAIBadge(res);
      populateAIModal(res);
    } catch (e) {
      console.error('Lỗi tải cấu hình AI:', e);
    }
  }

  function updateAIBadge(cfg) {
    const badge = document.getElementById('supp-ai-badge');
    if (!badge) return;
    if (cfg && cfg.is_active) {
      const pName = (cfg.effective_provider || cfg.provider) === 'gemini' ? 'Gemini' : 'OpenAI';
      badge.className = 'badge supp-ai-badge-active';
      badge.textContent = `AI: Có key (${pName})`;
      badge.title = 'Có cấu hình key; kết nối và kết quả đọc ảnh được kiểm tra khi sử dụng.';
    } else {
      badge.className = 'badge supp-ai-badge-inactive';
      badge.textContent = 'AI: Chưa thiết lập';
      badge.title = 'Nhấp vào "Cài đặt AI" để cấu hình API Key Google Gemini hoặc OpenAI';
    }
  }

  function populateAIModal(cfg) {
    if (!cfg) return;
    const provider = cfg.provider || 'gemini';
    const radio = document.querySelector(`input[name="ai_provider"][value="${provider}"]`);
    if (radio) radio.checked = true;
    switchAIProviderView(provider);

    // Gemini
    const geminiModel = document.getElementById('supp-gemini-model');
    if (geminiModel && cfg.gemini_model) geminiModel.value = cfg.gemini_model;
    const geminiKey = document.getElementById('supp-gemini-key');
    if (geminiKey) {
      geminiKey.value = '';
      geminiKey.placeholder = cfg.has_gemini_key ? `Đã lưu (${cfg.gemini_api_key}) - Để trống nếu không đổi` : 'Dán Gemini API Key tại đây (AIzaSy...)';
    }

    // OpenAI
    const openaiModel = document.getElementById('supp-openai-model');
    if (openaiModel && cfg.openai_model) openaiModel.value = cfg.openai_model;
    const openaiKey = document.getElementById('supp-openai-key');
    if (openaiKey) {
      openaiKey.value = '';
      openaiKey.placeholder = cfg.has_openai_key ? `Đã lưu (${cfg.openai_api_key}) - Để trống nếu không đổi` : 'Dán OpenAI API Key tại đây (sk-...)';
    }

    // Anti-AI Toggle
    const antiAiToggle = document.getElementById('supp-anti-ai-toggle');
    if (antiAiToggle) {
      antiAiToggle.checked = cfg.anti_ai_enabled !== false;
    }

    // Prompt
    const promptArea = document.getElementById('supp-prompt-template');
    if (promptArea) {
      promptArea.value = cfg.inspection_prompt || cfg.default_inspection_prompt || '';
    }
  }

  function switchAIProviderView(provider) {
    const cardGemini = document.getElementById('card-provider-gemini');
    const cardOpenai = document.getElementById('card-provider-openai');
    const secGemini = document.getElementById('supp-ai-gemini-section');
    const secOpenai = document.getElementById('supp-ai-openai-section');

    if (provider === 'gemini') {
      if (cardGemini) cardGemini.classList.add('active');
      if (cardOpenai) cardOpenai.classList.remove('active');
      if (secGemini) secGemini.style.display = 'block';
      if (secOpenai) secOpenai.style.display = 'none';
    } else {
      if (cardOpenai) cardOpenai.classList.add('active');
      if (cardGemini) cardGemini.classList.remove('active');
      if (secOpenai) secOpenai.style.display = 'block';
      if (secGemini) secGemini.style.display = 'none';
    }
  }

  function setupAISettings() {
    const btnOpen = document.getElementById('btn-open-ai-settings');
    const modal = document.getElementById('supp-ai-modal');
    const btnClose = document.getElementById('btn-close-ai-modal');
    const btnCancel = document.getElementById('btn-cancel-ai-modal');
    const btnSave = document.getElementById('btn-save-ai-settings');
    const btnTest = document.getElementById('btn-test-ai-connection');
    const btnRestorePrompt = document.getElementById('btn-restore-default-prompt');
    const testResultEl = document.getElementById('supp-ai-test-result');

    if (btnOpen && modal) {
      btnOpen.onclick = () => {
        if (aiConfig) populateAIModal(aiConfig);
        if (testResultEl) testResultEl.textContent = '';
        modal.style.display = 'flex';
      };
    }

    const closeModal = () => {
      if (modal) modal.style.display = 'none';
    };

    if (btnClose) btnClose.onclick = closeModal;
    if (btnCancel) btnCancel.onclick = closeModal;

    // Click outside modal dialog to close
    if (modal) {
      modal.onclick = (e) => {
        if (e.target === modal) closeModal();
      };
    }

    // Provider radio switch
    document.querySelectorAll('input[name="ai_provider"]').forEach(r => {
      r.onchange = () => {
        switchAIProviderView(r.value);
      };
    });

    // Eye toggle for password fields
    document.querySelectorAll('.supp-btn-toggle-eye').forEach(btn => {
      btn.onclick = () => {
        const targetId = btn.getAttribute('data-target');
        const input = document.getElementById(targetId);
        if (!input) return;
        if (input.type === 'password') {
          input.type = 'text';
          btn.textContent = '🔒';
        } else {
          input.type = 'password';
          btn.textContent = '👁️';
        }
      };
    });

    // Test connection
    if (btnTest) {
      btnTest.onclick = async () => {
        const provider = document.querySelector('input[name="ai_provider"]:checked')?.value || 'gemini';
        const model = provider === 'gemini' ?
          document.getElementById('supp-gemini-model')?.value :
          document.getElementById('supp-openai-model')?.value;
        const keyInput = provider === 'gemini' ?
          document.getElementById('supp-gemini-key') :
          document.getElementById('supp-openai-key');
        const apiKey = keyInput?.value.trim() || '';

        btnTest.disabled = true;
        btnTest.innerHTML = '<span>⏳ Đang kiểm tra...</span>';
        if (testResultEl) {
          testResultEl.textContent = 'Đang gửi ping kiểm tra...';
          testResultEl.style.color = '#38bdf8';
        }

        try {
          const res = await fetch('/api/supplement/ai-test', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ provider, model, api_key: apiKey })
          }).then(r => r.json());

          if (res.success) {
            if (testResultEl) {
              testResultEl.textContent = '✅ ' + res.message;
              testResultEl.style.color = '#4ade80';
            }
          } else {
            if (testResultEl) {
              testResultEl.textContent = '❌ ' + (res.message || 'Lỗi kết nối');
              testResultEl.style.color = '#f87171';
            }
          }
        } catch (e) {
          if (testResultEl) {
            testResultEl.textContent = '❌ Lỗi mạng: ' + e.message;
            testResultEl.style.color = '#f87171';
          }
        } finally {
          btnTest.disabled = false;
          btnTest.innerHTML = '<span>⚡ Kiểm tra kết nối API</span>';
        }
      };
    }

    // Restore default prompt
    if (btnRestorePrompt) {
      btnRestorePrompt.onclick = () => {
        if (aiConfig && aiConfig.default_inspection_prompt) {
          const promptArea = document.getElementById('supp-prompt-template');
          if (promptArea) promptArea.value = aiConfig.default_inspection_prompt;
        }
      };
    }

    // Save AI settings
    if (btnSave) {
      btnSave.onclick = async () => {
        const provider = document.querySelector('input[name="ai_provider"]:checked')?.value || 'gemini';
        const geminiModel = document.getElementById('supp-gemini-model')?.value;
        const geminiKey = document.getElementById('supp-gemini-key')?.value.trim();
        const openaiModel = document.getElementById('supp-openai-model')?.value;
        const openaiKey = document.getElementById('supp-openai-key')?.value.trim();
        const promptTemplate = document.getElementById('supp-prompt-template')?.value;
        const antiAiToggle = document.getElementById('supp-anti-ai-toggle');
        const antiAiEnabled = antiAiToggle ? antiAiToggle.checked : true;

        const payload = {
          provider,
          gemini_model: geminiModel,
          openai_model: openaiModel,
          anti_ai_enabled: antiAiEnabled,
          inspection_prompt: promptTemplate
        };
        if (geminiKey) payload.gemini_api_key = geminiKey;
        if (openaiKey) payload.openai_api_key = openaiKey;

        btnSave.disabled = true;
        btnSave.innerHTML = '<span>⏳ Đang lưu...</span>';

        try {
          const res = await fetch('/api/supplement/ai-config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          }).then(r => r.json());

          if (res.success) {
            aiConfig = res.config;
            updateAIBadge(res.config);
            notify('Đã lưu cài đặt AI và Prompt thành công!', 'success');
            closeModal();
          } else {
            notify('Lỗi: ' + (res.error || 'Không thể lưu'), 'error');
          }
        } catch (e) {
          notify('Lỗi kết nối khi lưu: ' + e.message, 'error');
        } finally {
          btnSave.disabled = false;
          btnSave.innerHTML = '<span>💾 Lưu Cài Đặt AI</span>';
        }
      };
    }
  }

  // --- INITIALIZE ---
  async function init() {
    setupEmployeeCombobox();
    setupRandomPortraitPicker();
    setupDragAndDrop();
    setupBatchTools();
    setupFormSubmit();
    setupStagingActions();
    setupAISettings();

    await loadProjects();
    await loadStaging();
    await loadAIConfig();

    // Compatibility hook for other tabs / tests
    window.supplementLoadRecords = loadStaging;
    window.showWatermarkConfirmation = showWatermarkConfirmation;
    window.regenerateStagingPhoto = regenerateStagingPhoto;
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
