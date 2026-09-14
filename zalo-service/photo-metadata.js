const vietnamDate = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Ho_Chi_Minh', year: 'numeric', month: '2-digit', day: '2-digit',
});

function validDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return false;
    const date = new Date(`${value}T00:00:00Z`);
    return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function normalizeSendDate(photo) {
    let value = photo.timestamp;
    if (typeof value === 'string' && /^\d+(\.\d+)?$/.test(value)) value = Number(value);
    if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
        value = new Date(value < 1e11 ? value * 1000 : value);
    } else if (typeof value === 'string' && /T.*(?:Z|[+-]\d{2}:\d{2})$/i.test(value)) {
        value = new Date(value);
    } else {
        value = null;
    }
    if (value && Number.isFinite(value.getTime())) {
        const parts = Object.fromEntries(vietnamDate.formatToParts(value).map(p => [p.type, p.value]));
        return `${parts.year}-${parts.month}-${parts.day}`;
    }
    if (['message', 'header', 'viewer'].includes(photo.dateSource) && validDate(photo.date)) return photo.date;
    return '';
}

function mergePhotos(photos) {
    const result = [];
    for (const photo of photos) {
        const existing = result.find(p => (p.id === photo.id && (p.messageId || p.date === photo.date)) || (
            p.url === photo.url && p.date === photo.date && (!p.messageId || !photo.messageId)
        ));
        if (!existing) result.push({...photo});
        else {
            if (!existing.messageId && photo.messageId) {
                existing.id = photo.id;
                existing.messageId = photo.messageId;
                existing.imageIndex = photo.imageIndex;
                existing.elementToken = photo.elementToken;
            }
            if (!existing.timestamp && photo.timestamp) existing.timestamp = photo.timestamp;
            if (!existing.dateSource && photo.dateSource) existing.dateSource = photo.dateSource;
        }
    }
    return result;
}

// Self-contained: serialized by Puppeteer into the page. No dates inferred from URLs,
// no date state survives a scan, and every image in an album gets its own identity.
function collectVisiblePhotos() {
    const media = document.querySelector('#innerScrollContainer');
    const nodes = new Set([
        ...(media ? media.querySelectorAll('img') : []),
        ...document.querySelectorAll('.chat-message img.zimg-el, .msg-item img.zimg-el, .img-center-box img.zimg-el'),
    ]);
    function headerDate(el) {
        if (!el) return '';
        const id = (el.id || '').match(/^(?:date|item)_(\d{2})(\d{2})(\d{4})(?:\D|$)/);
        if (id) return `${id[3]}-${id[2]}-${id[1]}`;
        if (!el.matches('.media-item-date-group, .chat-date, .date-header, [id^="date_"]')) return '';
        const text = el.textContent.trim();
        const full = text.match(/(?:^|\D)(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\D|$)/)
            || text.match(/Ngày\s+(\d{1,2})\s+Tháng\s+(\d{1,2})\s+(?:Năm\s+)?(\d{4})/i);
        return full ? `${full[3]}-${full[2].padStart(2, '0')}-${full[1].padStart(2, '0')}` : '';
    }
    function precedingHeader(start, boundary) {
        for (let el = start; el && el !== boundary; el = el.parentElement) {
            const own = headerDate(el);
            if (own) return own;
            for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling) {
                // A header without a year is a boundary too: never leak an older header across it.
                const header = prev.matches('.media-item-date-group, .chat-date, .date-header, [id^="date_"]')
                    ? prev : null;
                if (header) return headerDate(header);
            }
        }
        return '';
    }
    return Array.from(nodes).flatMap(img => {
        const url = img.currentSrc || img.src;
        if (!url || /avatar|sticker|emoji/.test(url) || !img.getClientRects().length) return [];
        const message = img.closest('[data-msg-id], [data-message-id]') || img.closest('.chat-message, .msg-item');
        const messageId = message?.dataset.msgId || message?.dataset.messageId || '';
        const imageIndex = message ? Array.from(message.querySelectorAll('img')).indexOf(img) : 0;
        const timestamp = message?.dataset.sendTime || message?.dataset.timestamp
            || message?.querySelector('time[datetime]')?.getAttribute('datetime') || '';
        const inMedia = media?.contains(img);
        const boundary = inMedia ? media : img.closest('#messageViewContainer, .chat-message-list, #chatViewContainer, .chat-content');
        // Without a known chat boundary we must not borrow unrelated dates from the page.
        const date = boundary ? precedingHeader(message || img, boundary) : '';
        // Preserve the exact tile when the same URL was sent on multiple days.
        const elementToken = img.dataset.zaloPhotoToken || (img.dataset.zaloPhotoToken = crypto.randomUUID());
        return [{id: messageId ? `${messageId}:${imageIndex}` : `url:${url}`, messageId, elementToken,
            imageIndex, url, timestamp, date, dateSource: date ? 'header' : '',
            source: inMedia ? 'media_store' : 'chat_view'}];
    });
}

module.exports = {validDate, normalizeSendDate, mergePhotos, collectVisiblePhotos};
