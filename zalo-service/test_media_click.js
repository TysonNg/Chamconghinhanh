const puppeteer = require('puppeteer-core');
const path = require('path');
const fs = require('fs');

(async () => {
    try {
        const profileDir = path.resolve('zalo-browser-profile');
        const lockPath = path.join(profileDir, 'lockfile');
        try { if (fs.existsSync(lockPath)) fs.unlinkSync(lockPath); } catch (e) {}

        const b = await puppeteer.launch({
            executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            userDataDir: profileDir,
            headless: true,
            args: ['--no-sandbox', '--window-size=1600,1000']
        });
        const page = (await b.pages())[0];
        await page.goto('https://chat.zalo.me', { waitUntil: 'networkidle2', timeout: 30000 });

        // Search and open group
        const searchInput = await page.$('#contact-search-input');
        await searchInput.click();
        await searchInput.type('Dự án Long An', { delay: 50 });
        await new Promise(r => setTimeout(r, 2000));
        await page.evaluate(() => {
            const results = document.querySelectorAll('.conv-item, .search-item');
            if (results.length > 0) results[0].click();
        });
        await new Promise(r => setTimeout(r, 2500));

        // Ensure right sidebar is open
        await page.evaluate(() => {
            const btn = document.querySelector("div[title='Thông tin hội thoại'], div[title='Thông tin nhóm']");
            if (btn && !btn.className.includes('focused')) {
                btn.click();
            }
        });
        await new Promise(r => setTimeout(r, 2000));

        // Find "Xem tất cả" inside right sidebar
        const clicked = await page.evaluate(() => {
            const sidebar = document.querySelector('.chat-right-menu, .conversation-info, .sidebar');
            const root = sidebar || document;
            const buttons = Array.from(root.querySelectorAll('*')).filter(el => {
                const txt = el.textContent ? el.textContent.trim() : '';
                return (txt === 'Xem tất cả' || txt === 'Xem tất cả >') && el.children.length === 0;
            });
            if (buttons.length > 0) {
                buttons[0].click();
                return { success: true, text: buttons[0].textContent.trim() };
            }
            return { success: false, buttonsCount: buttons.length };
        });
        console.log('Clicked "Xem tất cả":', clicked);

        await new Promise(r => setTimeout(r, 3000));
        await page.screenshot({ path: path.join(__dirname, 'debug_after_xem_tat_ca.png') });

        // Inspect what is now shown
        const mediaElements = await page.evaluate(() => {
            // Find all images or video items
            const imgs = Array.from(document.querySelectorAll("img, .zimg-el, [style*='background-image']")).map(el => ({
                tag: el.tagName,
                src: el.src || el.getAttribute('data-src') || el.style?.backgroundImage,
                className: el.className,
                parentClass: el.parentElement?.className,
                closestMsg: el.closest('.chat-message, .msg-item, .media-item')?.className
            })).filter(x => x.src && !x.src.includes('avatar') && !x.src.includes('sticker') && !x.src.includes('iconlike'));

            return {
                totalMedia: imgs.length,
                samples: imgs.slice(0, 15)
            };
        });

        console.log('Media elements found:\n', JSON.stringify(mediaElements, null, 2));

        await b.close();
    } catch (e) {
        console.error('Error:', e);
    }
})();
