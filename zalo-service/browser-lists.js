// These functions run inside the browser through Puppeteer; keep them self-contained.
function collectVisibleGroups() {
    const scope = document.querySelector('.contact-pages, #group-list, #groups, .group-list, .contact-group-list');
    const rows = (scope || document).querySelectorAll('.contact-item-v2-wrapper, .group-item, .group-list-item, .contact-item, [data-id^="group-item"], [data-id^="group_item"]' + (scope ? ', .conv-item' : ''));
    return Array.from(rows).filter(row => row.getClientRects().length).flatMap(row => {
        const title = row.querySelector('.name, .group-item-title__name, .group-name, .group-title, .contact-name, .conv-item-title__name, [class*="name"]');
        const name = (title?.textContent || row.getAttribute('title') || '').trim();
        let id = row.getAttribute('data-group-id') || row.getAttribute('data-id') || row.id;
        // Zalo's current contact rows expose itemId on their React component,
        // rather than a DOM attribute. Never fall back to the name or row index.
        let fiber = row[Object.keys(row).find(key => /^__react(?:Fiber|InternalInstance)/.test(key))];
        for (let depth = 0; !id && fiber && depth < 8; depth++, fiber = fiber.return) {
            const itemId = fiber.memoizedProps?.itemId;
            if (typeof itemId === 'string' && itemId) id = itemId;
        }
        if (!name || !id) return [];

        // Extract avatar URL from the group row's zavatar or img element
        const avatarImg = row.querySelector('.zavatar img.a-child, img[class*="avatar"], .zavatar img');
        const avatar = avatarImg?.src || '';

        // Extract member count from subtitle/description text (e.g., "6 thành viên")
        const descEl = row.querySelector('.contact-desc, .contact-item-v2__desc, .group-desc, [class*="desc"], [class*="subtitle"]');
        const allText = row.textContent || '';
        const memberMatch = allText.match(/(\d+)\s*(?:thành viên|members?|người)/i);
        const totalMember = memberMatch ? parseInt(memberMatch[1], 10) : 0;
        const description = (descEl?.textContent || '').trim();

        return [{id, name, avatar, totalMember, description}];
    });
}

function scrollList(mode, move = false) {
    const visible = el => el && el.getClientRects().length;
    const media = document.querySelector('#innerScrollContainer');
    const groupScope = document.querySelector('.contact-pages, #group-list, #groups, .group-list, .contact-group-list');
    const anchor = mode === 'groups'
        ? Array.from((groupScope || document).querySelectorAll('.contact-item-v2-wrapper, .group-item, .group-list-item, .contact-item, [data-id^="group-item"], [data-id^="group_item"]' + (groupScope ? ', .conv-item' : ''))).find(visible) || groupScope
        : media || document.querySelector('#messageViewContainer, .chat-message-list, #chatViewContainer, .chat-content');
    let container = anchor;
    while (container && container !== document.body) {
        const style = getComputedStyle(container);
        if (/auto|scroll/.test(style.overflowY) && container.clientHeight > 0) break;
        container = container.parentElement;
    }
    if (container === document.body) container = null;
    const backwards = mode === 'photos' && !media;
    if (move && container) {
        // Overlap screens: advancing farther than the viewport skips virtualized rows.
        const step = Math.max(1, Math.floor(container.clientHeight * 0.75));
        if (move === 'start') container.scrollTop = backwards ? container.scrollHeight : 0;
        else container.scrollTop += backwards ? -step : step;
        container.dispatchEvent(new Event('scroll', {bubbles: true}));
    }
    const top = container?.scrollTop || 0;
    const height = container?.scrollHeight || 0;
    const viewport = container?.clientHeight || 0;
    const root = container || anchor?.parentElement;
    const loading = Array.from((root || document).querySelectorAll('[aria-busy="true"], .loading, .spinner, [class*="loading-indicator"]')).some(visible);
    const groupHeading = mode === 'groups' && Array.from(document.querySelectorAll('*')).find(el =>
        visible(el) && el.children.length === 0 && /(?:Danh sách nhóm|Nhóm)(?: và cộng đồng)?\s*\(\d+\)/i.test(el.textContent));
    const expectedCount = groupHeading ? Number(groupHeading.textContent.match(/\((\d+)\)/)[1]) : null;
    const empty = mode === 'groups' && /chưa (?:tham gia|có) nhóm|không có nhóm|no groups/i.test((groupScope || document.body).innerText);
    const notice = mode === 'photos'
        ? Array.from(document.querySelectorAll('#footer, .tds-media-list__footer-wrapper, .tds-media-list__footer-content')).filter(visible).map(el => el.textContent.trim()).join(' ')
        : '';
    return {top, height, viewport, loading, notice, expectedCount, empty,
        modeValid: mode !== 'photos' || !!(media && visible(media)),
        atEnd: !container || (backwards ? top <= 1 : top + viewport >= height - 2)};
}

function clickNavigation(labels) {
    const candidates = Array.from(document.querySelectorAll('button, [role="button"], [title], [data-id], a, span, div'));
    const element = candidates.find(el => el.getClientRects().length && labels.some(label =>
        el.getAttribute('title') === label || el.getAttribute('aria-label') === label ||
        (el.children.length === 0 && (el.textContent || '').trim().replace(/\s*\(\d+\)$/, '') === label)));
    if (!element) return false;
    element.click();
    return true;
}
function clickConversation(groupName, groupId) {
    // Broader selectors: search results on newer Zalo Web may use different class names
    const selectors = '.conv-item, .search-item, [data-id*="conv_item"], [id^="group-item-"], [id^="conv-item-"], [class*="search-result"], [class*="recent-item"], [class*="contact-item"]';
    const rows = [...new Set(Array.from(document.querySelectorAll(selectors))
        .map(el => el.matches('.conv-item, .search-item, [class*="contact-item"]') ? el : el.querySelector('.conv-item, .search-item') || el))].filter(el => el.getClientRects().length);

    function identities(el) {
        const rawList = [el.getAttribute('data-group-id'), el.getAttribute('data-id'), el.id].filter(Boolean);
        let fiber = el[Object.keys(el).find(key => /^__react(?:Fiber|InternalInstance)/.test(key))];
        for (let depth = 0; fiber && depth < 8; depth++, fiber = fiber.return) {
            for (const id of [fiber.memoizedProps?.id, fiber.memoizedProps?.itemId]) {
                if (typeof id === 'string' && id) rawList.push(id);
            }
        }
        const ids = new Set();
        for (const raw of rawList) {
            ids.add(raw);
            const clean = raw.replace(/^(?:group|conv)[-_]item[-_]/i, '');
            ids.add(clean);
            if (clean.startsWith('g')) {
                ids.add(clean.slice(1));
            } else if (/^\d+$/.test(clean)) {
                ids.add('g' + clean);
            }
        }
        return Array.from(ids);
    }

    const targetIds = groupId ? [
        groupId,
        groupId.replace(/^g/, ''),
        groupId.startsWith('g') ? groupId : 'g' + groupId,
        `group-item-${groupId}`,
        `group-item-${groupId.replace(/^g/, '')}`
    ] : [];

    const byId = targetIds.length && rows.find(el => {
        const ids = identities(el);
        return targetIds.some(tid => ids.includes(tid));
    });

    function dispatchClick(targetEl) {
        targetEl.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        targetEl.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
        targetEl.click();
    }

    if (byId) {
        dispatchClick(byId);
        return true;
    }

    const nameNorm = (groupName || '').trim().normalize('NFC').toLocaleLowerCase('vi');
    const target = nameNorm.replace(/\s+/g, ' ');

    const matches = rows.filter(el => {
        const ids = identities(el);
        if (targetIds.length && ids.length && !targetIds.some(tid => ids.includes(tid))) return false;

        const title = el.querySelector('.conv-item-title__name, .group-name, .contact-name, .search-item-name, .name, .truncate, [class*="title__name"], [class*="name"]');
        const text = (title?.textContent || el.getAttribute('title') || '').trim().normalize('NFC').toLocaleLowerCase('vi');
        const norm = text.replace(/\s+/g, ' ');

        // Exact match on title text
        if (norm === target) return true;
        // Truncated match with ellipsis (Zalo sometimes truncates long names in search results)
        if ((norm.endsWith('...') || norm.endsWith('…')) && target.startsWith(norm.replace(/(\.\.\.|…)$/, '').trim())) return true;

        // Fallback: check full text content of the row (stripping member counts or prefixes)
        const full = (el.textContent || '').trim().normalize('NFC').toLocaleLowerCase('vi');
        const stripped = full.replace(/^\s*\d+\s*/, '').replace(/\s+/g, ' ');
        if (stripped === target || stripped.startsWith(target)) return true;

        return false;
    });

    if (matches.length > 1) {
        // If multiple matches by name, see if any matches the targetId
        const preferred = targetIds.length && matches.find(el => {
            const ids = identities(el);
            return targetIds.some(tid => ids.includes(tid));
        });
        if (preferred) {
            dispatchClick(preferred);
            return true;
        }
        throw new Error(`Có nhiều nhóm tên "${groupName}"; không xác định được nhóm theo ID.`);
    }

    if (!matches.length) return false;
    dispatchClick(matches[0]);
    return true;
}

/**
 * Click the "Nhóm" (Groups) or "Liên hệ" (Contacts/Groups) tab in search results to filter results.
 * Returns true if a matching tab was found and clicked.
 */
function clickSearchGroupTab() {
    const tabRegexes = [
        /^(?:Nhóm|Groups?)(?:\s*(?:&|và)\s*cộng\s*đồng)?(?:\s*\(\d+\))?$/i,
        /^(?:Liên\s*hệ|Contacts?)(?:\s*\(\d+\))?$/i
    ];
    const tabs = Array.from(document.querySelectorAll('button, [role="tab"], .tab-item, a, span, div')).filter(el => {
        if (!el.getClientRects().length) return false;
        const text = (el.textContent || '').trim();
        return tabRegexes.some(rx => rx.test(text));
    });
    // Pick the most specific tab element (button or role="tab" or class includes tab)
    const target = tabs.find(t => t.tagName === 'BUTTON' || t.getAttribute('role') === 'tab' || (t.className && String(t.className).includes('tab'))) || tabs[0];
    if (target) {
        target.click();
        return true;
    }
    return false;
}

/**
 * Extract debug info about the current search results DOM (used when search fails).
 */
function debugSearchDom(groupName) {
    const allVisible = Array.from(document.querySelectorAll('.conv-item, .search-item, [data-id*="conv_item"], [class*="search-result"], [class*="recent-item"], [class*="contact-item"]')).filter(el => el.getClientRects().length);
    const items = allVisible.slice(0, 10).map(el => {
        const title = el.querySelector('.conv-item-title__name, .group-name, .contact-name, .search-item-name, .truncate, [class*="title__name"]');
        return {
            tag: el.tagName,
            className: (el.className || '').slice(0, 100),
            titleText: (title?.textContent || '').trim().slice(0, 80),
            fullText: (el.textContent || '').trim().slice(0, 120),
        };
    });
    const tabs = Array.from(document.querySelectorAll('button, [role="tab"]')).filter(el => el.getClientRects().length).map(el => ({
        text: (el.textContent || '').trim(),
        className: (el.className || '').slice(0, 60),
        active: el.className?.includes('active') || el.getAttribute('aria-selected') === 'true'
    }));
    const searchInput = document.querySelector('#contact-search-input, input[placeholder*="Tìm kiếm"]');
    return {
        query: groupName,
        searchInputValue: searchInput?.value || '',
        totalVisibleItems: allVisible.length,
        items,
        tabs: tabs.slice(0, 8),
        url: window.location.href,
        bodyTextSample: document.body.innerText.slice(0, 300)
    };
}
module.exports = {collectVisibleGroups, scrollList, clickNavigation, clickConversation, clickSearchGroupTab, debugSearchDom};
