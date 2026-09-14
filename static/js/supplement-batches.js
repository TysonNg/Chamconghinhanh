/* Supplement dates are record annotations, never capture timestamps. */
(() => {
  'use strict';
  const el = (tag, text) => { const n = document.createElement(tag); if (text) n.textContent = text; return n; };
  const message = text => { document.getElementById('batch-message').textContent = text; };
  let dirty = false;
  async function api(path, options) {
    const response = await fetch('/api/supplement/batches' + path, options);
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(data.error || `Yêu cầu thất bại (${response.status})`);
    return data;
  }
  const json = (method, data) => ({method, headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)});
  async function run(button, task) {
    button.disabled = true;
    try { await task(); message('Đã lưu.'); } catch (e) { message(e.message); }
    finally { button.disabled = false; }
  }
  async function list() {
    if (dirty) { message('Hãy lưu dòng đang sửa trước khi đổi batch.'); return; }
    try {
      const data = await api('');
      const host = document.getElementById('batch-list'); host.replaceChildren();
      for (const b of data.batches) {
        const button = el('button', `${b.employee} · ${b.items.length} ảnh · ${b.status === 'approved' ? 'Đã duyệt' : 'Bản nháp'}`);
        button.className = 'btn btn-secondary';
        button.onclick = () => run(button, async () => {
          if (dirty) throw new Error('Hãy lưu dòng đang sửa trước khi đổi batch.');
          show((await api('/' + b.id)).batch);
        });
        host.append(button);
      }
    } catch (e) { message(e.message); }
  }
  function show(b) {
    dirty = false;
    const host = document.getElementById('batch-detail'); host.replaceChildren();
    host.append(el('h3', b.employee), el('p', b.label));
    const approve = el('button', 'Duyệt hồ sơ'); approve.className = 'btn btn-primary';
    const download = el('a', 'Tải ZIP ảnh gốc + manifest');
    download.href = `/api/supplement/batches/${b.id}/download`;
    download.hidden = b.status !== 'approved';
    const rows = [];
    for (const item of b.items) {
      const row = el('div'); row.className = 'card'; row.style.padding = '12px';
      const link = el('a'); link.href = `/api/supplement/batches/${b.id}/items/${item.id}/image`;
      link.target = '_blank'; link.rel = 'noopener';
      const img = el('img'); img.src = link.href; img.alt = item.source_name;
      img.loading = 'lazy'; img.style.cssText = 'width:160px;height:120px;object-fit:contain'; link.append(img);
      const dayLabel = el('label', 'Ngày hồ sơ'); const day = el('input'); day.type = 'date'; day.value = item.supplement_date; dayLabel.append(day);
      const noteLabel = el('label', 'Ghi chú'); const note = el('input'); note.value = item.note; note.maxLength = 2000; noteLabel.append(note);
      const state = {item, day, note, changed: false}; rows.push(state);
      const change = () => { state.changed = true; dirty = true; approve.disabled = true; download.hidden = true; };
      day.oninput = note.oninput = change;
      row.append(link, el('p', item.source_name), dayLabel, noteLabel); host.append(row);
    }
    const save = el('button', 'Lưu các dòng đã sửa'); save.className = 'btn btn-secondary';
    save.onclick = () => run(save, async () => {
      for (const r of rows.filter(r => r.changed)) {
        await api(`/${b.id}/items/${r.item.id}`, json('PATCH', {supplement_date: r.day.value, note: r.note.value}));
        r.changed = false;
      }
      show((await api('/' + b.id)).batch); await list();
    });
    const reviewerLabel = el('label', 'Tên người duyệt (tự khai, chưa xác thực tài khoản)');
    const reviewer = el('input'); reviewer.maxLength = 200; reviewerLabel.append(reviewer);
    approve.onclick = () => run(approve, async () => {
      if (dirty) throw new Error('Hãy lưu các dòng đã sửa trước khi duyệt.');
      show((await api('/' + b.id + '/approve', json('POST', {reviewer: reviewer.value}))).batch); await list();
    });
    const history = el('details'); history.append(el('summary', 'Lịch sử hồ sơ'), el('pre', JSON.stringify(b.history, null, 2)));
    host.append(save, reviewerLabel, approve, download, history);
  }
  document.getElementById('batch-create').onsubmit = event => {
    event.preventDefault();
    const form = event.currentTarget;
    run(form.querySelector('button'), async () => {
      if (dirty) throw new Error('Hãy lưu batch đang sửa trước.');
      const files = form.elements.photos.files;
      if (!files.length || files.length > 31) throw new Error('Chọn từ 1 đến 31 ảnh.');
      show((await api('', {method: 'POST', body: new FormData(form)})).batch); await list();
    });
  };
  document.getElementById('batch-refresh').onclick = list;
  window.supplementLoadRecords = list;
  window.addEventListener('beforeunload', e => { if (dirty) { e.preventDefault(); e.returnValue = ''; } });
})();
