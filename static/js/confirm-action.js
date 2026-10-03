/* Shared asynchronous confirmation; duplicate prompts are cancelled. */
(() => {
  'use strict';
  let pending = false;
  window.confirmAction = function confirmAction(message) {
    if (pending) return Promise.resolve(false);
    pending = true;
    return new Promise(resolve => {
      const previous = document.activeElement;
      const overflow = document.body.style.overflow;
      const backdrop = document.createElement('div');
      backdrop.id = 'confirm-action-backdrop';
      backdrop.style.cssText = 'position:fixed;inset:0;z-index:100000;background:#0008;display:flex;align-items:center;justify-content:center;padding:20px';
      const dialog = document.createElement('section');
      dialog.setAttribute('role', 'alertdialog');
      dialog.setAttribute('aria-modal', 'true');
      dialog.setAttribute('aria-labelledby', 'confirm-action-title');
      dialog.setAttribute('aria-describedby', 'confirm-action-message');
      dialog.style.cssText = 'background:var(--bg-primary,#fff);color:var(--text-primary,#222);padding:24px;border-radius:12px;width:440px;max-width:100%;max-height:80vh;overflow:auto;box-sizing:border-box';
      const title = document.createElement('h3'); title.id = 'confirm-action-title'; title.textContent = 'Xác nhận thao tác';
      const text = document.createElement('p'); text.id = 'confirm-action-message'; text.textContent = message; text.style.whiteSpace = 'pre-wrap';
      const cancel = document.createElement('button'); cancel.type = 'button'; cancel.textContent = 'Hủy bỏ'; cancel.className = 'btn btn-secondary'; cancel.dataset.confirm = 'cancel';
      const accept = document.createElement('button'); accept.type = 'button'; accept.textContent = 'Xác nhận'; accept.className = 'btn btn-danger'; accept.dataset.confirm = 'accept'; accept.style.marginLeft = '12px';
      let finished = false;
      function finish(value) {
        if (finished) return; finished = true;
        document.removeEventListener('keydown', keydown, true);
        document.removeEventListener('focusin', focusin, true);
        backdrop.remove(); document.body.style.overflow = overflow; pending = false;
        if (previous && previous.isConnected) previous.focus();
        resolve(value);
      }
      function focusin(event) { if (!dialog.contains(event.target)) cancel.focus(); }
      function keydown(event) {
        if (event.key === 'Escape') { event.preventDefault(); event.stopImmediatePropagation(); finish(false); }
        if (event.key === 'Tab') { event.preventDefault(); (document.activeElement === cancel ? accept : cancel).focus(); }
      }
      cancel.onclick = () => finish(false); accept.onclick = () => finish(true);
      backdrop.onclick = event => { if (event.target === backdrop) finish(false); };
      dialog.append(title, text, cancel, accept); backdrop.append(dialog); document.body.append(backdrop);
      document.body.style.overflow = 'hidden';
      document.addEventListener('keydown', keydown, true); document.addEventListener('focusin', focusin, true); cancel.focus();
    });
  };
})();
