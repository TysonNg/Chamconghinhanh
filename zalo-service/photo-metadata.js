const vietnamDate = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Ho_Chi_Minh', year: 'numeric', month: '2-digit', day: '2-digit',
});

function validDate(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '')) return false;
    const date = new Date(`${value}T00:00:00Z`);
    return Number.isFinite(date.getTime()) && date.toISOString().slice(0, 10) === value;
}

function normalizeMessageTimestamp(value) {
    const timestamp = typeof value === 'number' ? value
        : (typeof value === 'string' && /^\d+(?:\.\d+)?$/.test(value.trim()) ? Number(value) : NaN);
    return Number.isFinite(timestamp) && timestamp > 0 ? timestamp : '';
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

// Classify only a verified message timestamp; a day header has no shift time.
function classifySendShift(photo) {
    let value = photo.timestamp;
    if (typeof value === 'string' && /^\d+(\.\d+)?$/.test(value)) value = Number(value);
    let date;
    if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
        date = new Date(value < 1e11 ? value * 1000 : value);
    } else if (typeof value === 'string' && /T.*(?:Z|[+-]\d{2}:\d{2})$/i.test(value)) {
        date = new Date(value);
    }
    if (!date || !Number.isFinite(date.getTime())) return {shift: 'unknown', send_time: null};
    const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
        timeZone: 'Asia/Ho_Chi_Minh', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
    }).formatToParts(date).map(p => [p.type, p.value]));
    const seconds = Number(parts.hour) * 3600 + Number(parts.minute) * 60 + Number(parts.second);
    const shift = seconds >= 5 * 3600 && seconds <= 15 * 3600 ? 'morning'
        : seconds >= 16 * 3600 ? 'afternoon' : 'outside';
    return {shift, send_time: `${parts.hour}:${parts.minute}:${parts.second}`};
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
    const mediaImgs = media ? Array.from(media.querySelectorAll('img')) : [];
    // Khi Kho Media Store mở và có ảnh, chỉ quét ảnh trong Media Store để tránh ảnh chat nền
    // không có header ngày làm nhiễu danh sách.
    const nodes = media
        ? new Set(mediaImgs)
        : new Set([
            ...(media ? mediaImgs : []),
            ...document.querySelectorAll('.chat-message img.zimg-el, .msg-item img.zimg-el, .img-center-box img.zimg-el'),
        ]);
    function headerDate(el) {
        if (!el) return '';
        const id = (el.id || '').match(/^(?:date|item)_(\d{2})(\d{2})(\d{4})(?:\D|$)/);
        if (id) return `${id[3]}-${id[2]}-${id[1]}`;
        // Thêm nhiều selectors hơn cho date header (Zalo Web thay đổi thường xuyên)
        const isDateElement = el.matches('.media-item-date-group, .chat-date, .date-header, [id^="date_"], [id^="item_"], .date-group, [class*="date-group"], [class*="dateGroup"], [class*="time-divider"]');
        const text = el.textContent.trim();
        // Chấp nhận bất kỳ element nào có text chứa pattern ngày (không chỉ date selectors)
        if (!isDateElement && text.length > 30) return '';
        // Match DD/MM/YYYY
        const full = text.match(/(?:^|\D)(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\D|$)/)
            // Match "Ngày DD Tháng MM Năm YYYY"
            || text.match(/Ngày\s+(\d{1,2})\s+Tháng\s+(\d{1,2})\s+(?:Năm\s+)?(\d{4})/i)
            || text.match(/(\d{1,2})-(\d{1,2})-(\d{4})/)
            || text.match(/(\d{1,2})\.(\d{1,2})\.(\d{4})/);
        if (full) return `${full[3]}-${full[2].padStart(2, '0')}-${full[1].padStart(2, '0')}`;
        // *** CRITICAL FIX: Zalo Media Store hiển thị "Ngày DD Tháng M" KHÔNG CÓ NĂM ***
        const noYear = text.match(/Ngày\s+(\d{1,2})\s+Tháng\s+(\d{1,2})\s*$/i);
        if (noYear) {
            const year = new Date().getFullYear();
            return `${year}-${noYear[2].padStart(2, '0')}-${noYear[1].padStart(2, '0')}`;
        }
        return '';
    }
    function precedingHeader(start, boundary) {
        for (let el = start; el && el !== boundary; el = el.parentElement) {
            const own = headerDate(el);
            if (own) return own;
            for (let prev = el.previousElementSibling; prev; prev = prev.previousElementSibling) {
                // A header without a year is a boundary too: never leak an older header across it.
                const header = prev.matches('.media-item-date-group, .chat-date, .date-header, [id^="date_"], [id^="item_"], .date-group, [class*="date-group"], [class*="dateGroup"], [class*="time-divider"]')
                    ? prev : null;
                if (header) return headerDate(header);
                // Fallback: Zalo Media Store có thể dùng element thường cho date header
                // Chỉ thử nếu element có ít text (giống header, không phải container ảnh)
                const prevText = (prev.textContent || '').trim();
                if (prevText.length > 0 && prevText.length < 40 && prev.querySelectorAll('img').length === 0) {
                    const prevDate = headerDate(prev);
                    if (prevDate) return prevDate;
                }
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

        // Trích xuất timestamp từ ID của img (dạng img-1790160181783.5914458...)
        const imgIdMatch = (img.id || '').match(/img-(\d{12,14})\./);
        const idTimestamp = imgIdMatch ? Number(imgIdMatch[1]) : '';

        // Trích xuất ngày từ container dòng item (dạng item_23092026_...)
        const itemRow = img.closest('[id^="item_"]');
        const rowDate = itemRow ? headerDate(itemRow) : '';

        // Mở rộng nguồn lấy timestamp: thêm data-ts, data-time, title attribute, imgId
        const timestamp = message?.dataset.sendTime || message?.dataset.timestamp || message?.dataset.ts || message?.dataset.time
            || message?.querySelector('time[datetime]')?.getAttribute('datetime')
            || img.closest('[data-send-time]')?.dataset.sendTime
            || img.closest('[data-timestamp]')?.dataset.timestamp
            || idTimestamp || '';

        const inMedia = media?.contains(img);
        const boundary = inMedia ? media : img.closest('#messageViewContainer, .chat-message-list, #chatViewContainer, .chat-content');

        // Without a known chat boundary we must not borrow unrelated dates from the page.
        let date = rowDate || (boundary ? precedingHeader(message || img, boundary) : '');
        // Fallback: thử lấy ngày từ container cha gần nhất có class chứa date info
        if (!date && inMedia) {
            const parentGroup = img.closest('[class*="date-group"], [class*="dateGroup"], [id^="date_"], [id^="item_"]');
            if (parentGroup) date = headerDate(parentGroup);
        }
        // Preserve the exact tile when the same URL was sent on multiple days.
        const randToken = (window.crypto && typeof window.crypto.randomUUID === 'function')
            ? window.crypto.randomUUID()
            : (Date.now().toString(36) + Math.random().toString(36).slice(2));
        const elementToken = img.dataset.zaloPhotoToken || (img.dataset.zaloPhotoToken = randToken);
        return [{id: messageId ? `${messageId}:${imageIndex}` : `url:${url}`, messageId, elementToken,
            imageIndex, url, timestamp, date, dateSource: date ? 'header' : (timestamp ? 'message' : ''),
            source: inMedia ? 'media_store' : 'chat_view'}];
    });
}

module.exports = {classifySendShift, validDate, normalizeMessageTimestamp, normalizeSendDate, mergePhotos, collectVisiblePhotos};
