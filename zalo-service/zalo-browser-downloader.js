const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer-core");
const {validDate, normalizeSendDate, mergePhotos, collectVisiblePhotos} = require("./photo-metadata");
const {readImage, resolvePhoto} = require("./photo-viewer");

function getBrowserExecutablePath() {
    const candidatePaths = [
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
        "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
    ];
    for (const p of candidatePaths) {
        if (fs.existsSync(p)) return p;
    }
    throw new Error("Không tìm thấy Google Chrome hoặc Microsoft Edge trên máy tính của bạn.");
}

class ZaloBrowserDownloader {
    constructor(projectRootDir) {
        this.projectRootDir = projectRootDir || path.resolve(__dirname, "..");
        this.inputImagesDir = path.join(this.projectRootDir, "input_images");
        this.profileDir = path.join(__dirname, "zalo-browser-profile");
        
        if (!fs.existsSync(this.profileDir)) {
            fs.mkdirSync(this.profileDir, { recursive: true });
        }

        this.browser = null;
        this.page = null;
        this.isLoggedIn = false;
        this.isDownloading = false;
        this.currentQrBase64 = null;

        this.progress = {
            status: "idle", // idle | waiting_qr | opening_group | scanning_media | downloading | done | error
            current: 0,
            total: 0,
            downloaded: 0,
            skipped: 0,
            failed: 0,
            lowQuality: 0,
            unknownDate: 0,
            filteredOut: 0,
            counterVersion: 2,
            currentFile: "",
            log: [],
            error: null,
            targetFolder: null
        };
    }

    _log(message) {
        const time = new Date().toLocaleTimeString("vi-VN", { hour12: false });
        const entry = `[${time}] ${message}`;
        console.log(`[BrowserDownloader] ${entry}`);
        this.progress.log.push(entry);
        if (this.progress.log.length > 200) {
            this.progress.log.shift();
            this.progress.logOffset = (this.progress.logOffset || 0) + 1;
        }
    }

    getStatus() {
        return {
            isLoggedIn: this.isLoggedIn,
            isDownloading: this.isDownloading,
            currentQrBase64: this.currentQrBase64,
            progress: this.progress
        };
    }

    _cleanupStaleLocks() {
        const lockPath = path.join(this.profileDir, "lockfile");
        if (fs.existsSync(lockPath)) {
            try {
                fs.unlinkSync(lockPath);
                this._log("Đã dọn dẹp file lock profile cũ.");
            } catch (e) {
                try {
                    const { execSync } = require("child_process");
                    const psCmd = `powershell.exe -NoProfile -Command "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -match 'zalo-browser-profile' } | ForEach-Object { $_.ProcessId }"`;
                    const out = execSync(psCmd, { encoding: "utf8" });
                    const pids = out.split(/\r?\n/).map(s => s.trim()).filter(s => /^\d+$/.test(s));
                    for (const pid of pids) {
                        try { execSync(`taskkill /F /PID ${pid}`); } catch {}
                    }
                } catch {}
                try {
                    if (fs.existsSync(lockPath)) fs.unlinkSync(lockPath);
                } catch {}
            }
        }
    }

    resolvePortraitDir() {
        const candidates = ["Ảnh BV", "Anh BV", "ANH BV", "anh bv", "ẢnhBV"];
        for (const c of candidates) {
            const p = path.join(this.projectRootDir, c);
            if (fs.existsSync(p)) return p;
        }
        return path.join(this.projectRootDir, "Ảnh BV");
    }

    ensureProjectStructure(projectName) {
        if (!projectName) return;
        const safeName = projectName.replace(/[<>:"/\\|?*]/g, '_').trim();
        if (!safeName) return;

        // 1. Thư mục ảnh camera theo ngày trong input_images/<projectName>
        const inputProjDir = path.join(this.inputImagesDir, safeName);
        if (!fs.existsSync(inputProjDir)) {
            fs.mkdirSync(inputProjDir, { recursive: true });
        }
        for (let d = 1; d <= 31; d++) {
            const dayDir = path.join(inputProjDir, String(d).padStart(2, '0'));
            if (!fs.existsSync(dayDir)) {
                fs.mkdirSync(dayDir, { recursive: true });
            }
        }

        // 2. Thư mục chân dung nhân viên trong Ảnh BV/<projectName>
        const portraitDir = this.resolvePortraitDir();
        const portraitProjDir = path.join(portraitDir, safeName);
        if (!fs.existsSync(portraitProjDir)) {
            fs.mkdirSync(portraitProjDir, { recursive: true });
        }
    }

    /**
     * Launch or connect to browser
     */
    async getOrLaunchBrowser(headless = true) {
        if (this.browser && this.browser.connected) {
            // Nếu chế độ headless yêu cầu khác với trạng thái hiện tại, khởi động lại để đổi chế độ hiển thị
            if (this.currentHeadless === headless) {
                return this.browser;
            }
            this._log(`Chuyển đổi chế độ trình duyệt sang: ${headless ? "Chạy ngầm" : "Hiển thị cửa sổ"}`);
            await this.close();
        }

        this._cleanupStaleLocks();

        const exePath = getBrowserExecutablePath();
        this._log(`Khởi chạy trình duyệt (${headless ? "Chạy ngầm" : "Hiển thị cửa sổ"}): ${path.basename(exePath)}`);

        const browserArgs = [
            "--no-sandbox",
            "--disable-setuid-sandbox",
            "--disable-dev-shm-usage",
            "--disable-blink-features=AutomationControlled",
            "--window-size=1280,850",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-backgrounding-occluded-windows",
            "--disable-renderer-backgrounding"
        ];
        if (!headless) {
            browserArgs.push("--start-maximized");
        }

        this.browser = await puppeteer.launch({
            executablePath: exePath,
            userDataDir: this.profileDir,
            headless: headless ? true : false,
            defaultViewport: null,
            args: browserArgs
        });

        this.currentHeadless = headless;

        this.browser.on("disconnected", () => {
            this._log("Trình duyệt đã đóng.");
            this.browser = null;
            this.page = null;
        });

        const pages = await this.browser.pages();
        this.page = pages.length > 0 ? pages[0] : await this.browser.newPage();
        await this.page.setViewport({ width: 1280, height: 850 });

        if (!headless) {
            try { await this.page.bringToFront(); } catch {}
        }

        return this.browser;
    }

    /**
     * Ensure user is logged in to Zalo Web
     */
    async ensureLoggedIn(headless = true) {
        await this.getOrLaunchBrowser(headless);

        const currentUrl = this.page.url();
        if (!currentUrl.includes("chat.zalo.me") || currentUrl.includes("id.zalo.me")) {
            this._log("Đang mở trang https://chat.zalo.me...");
            await this.page.goto("https://chat.zalo.me", { waitUntil: "domcontentloaded", timeout: 45000 });
        }

        if (!headless) {
            try {
                await this.page.bringToFront();
                const { exec } = require("child_process");
                const psFocus = `powershell -NoProfile -Command "$wshell = New-Object -ComObject wscript.shell; $procs = Get-Process -Name chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -match 'Zalo' }; if ($procs) { $wshell.AppActivate($procs[0].Id) }"`;
                exec(psFocus);
            } catch (e) {}
        }

        this._log("Đang kiểm tra trạng thái đăng nhập Zalo Web...");

        let state = null; // 'logged_in' | 'needs_login'
        const startCheck = Date.now();
        while (Date.now() - startCheck < 45000) {
            const url = this.page.url();
            if (url.includes("id.zalo.me") || url.includes("/login")) {
                state = "needs_login";
                break;
            }

            const isQrVisible = await this.page.$(".qr-container, .qrcode, canvas").catch(() => null);
            if (isQrVisible) {
                state = "needs_login";
                break;
            }

            const isChatReady = await this.page.$("#contact-search-input, .conv-item, #main-tab").catch(() => null);
            if (isChatReady) {
                state = "logged_in";
                break;
            }

            await new Promise(r => setTimeout(r, 1000));
        }

        // If state is still unknown, try reloading once
        if (state === null) {
            this._log("Trang Zalo Web tải chậm, đang thử tải lại trang...");
            try {
                await this.page.reload({ waitUntil: "domcontentloaded", timeout: 30000 });
                await new Promise(r => setTimeout(r, 5000));

                const url2 = this.page.url();
                if (url2.includes("id.zalo.me") || url2.includes("/login")) {
                    state = "needs_login";
                } else {
                    const isQr2 = await this.page.$(".qr-container, .qrcode, canvas").catch(() => null);
                    if (isQr2) {
                        state = "needs_login";
                    } else {
                        const isChat2 = await this.page.$("#contact-search-input, .conv-item, #main-tab").catch(() => null);
                        if (isChat2) {
                            state = "logged_in";
                        }
                    }
                }
            } catch (reloadErr) {
                this._log(`Không thể reload trang: ${reloadErr.message}`);
            }
        }

        if (state === "needs_login" || this.page.url().includes("id.zalo.me")) {
            this.isLoggedIn = false;
            this.progress.status = "waiting_qr";

            // Nếu đang chạy headless mà chưa đăng nhập, tự động mở cửa sổ Chrome để người dùng quét mã trực tiếp
            if (this.currentHeadless === true) {
                this._log("Đang mở cửa sổ Chrome hiển thị mã QR để bạn quét trực tiếp trên màn hình...");
                await this.close();
                await this.getOrLaunchBrowser(false);
                await this.page.goto("https://chat.zalo.me", { waitUntil: "domcontentloaded", timeout: 45000 });
                await new Promise(r => setTimeout(r, 3000));
            }

            if (!headless || this.currentHeadless === false) {
                try {
                    await this.page.bringToFront();
                    const { exec } = require("child_process");
                    const psFocus = `powershell -NoProfile -Command "$wshell = New-Object -ComObject wscript.shell; $procs = Get-Process -Name chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -match 'Zalo' }; if ($procs) { $wshell.AppActivate($procs[0].Id) }"`;
                    exec(psFocus);
                } catch (e) {}
            }

            this._log("👉 Vui lòng mở app Zalo trên điện thoại, quét mã QR trên màn hình Chrome hoặc ngay trên giao diện web để đăng nhập.");

            // Chụp ảnh QR và đẩy trực tiếp lên giao diện web
            try {
                await this.page.waitForSelector(".qr-container, .qrcode, canvas", { timeout: 8000 });
                const qrElement = await this.page.$(".qr-container, .qrcode, canvas");
                if (qrElement) {
                    const b64 = await qrElement.screenshot({ encoding: "base64" });
                    this.currentQrBase64 = "data:image/png;base64," + b64;
                    this.progress.qrImage = this.currentQrBase64;
                    try {
                        fs.writeFileSync(path.join(__dirname, "qr.png"), Buffer.from(b64, "base64"));
                    } catch {}
                    this._log("Đã tải mã QR hiển thị ngay trên giao diện web để bạn quét tiện lợi.");
                }
            } catch (e) {}

            // Chờ người dùng quét mã trên điện thoại (tối đa 120s)
            this._log("Đang chờ xác nhận quét mã trên điện thoại (tối đa 2 phút)...");
            const startWait = Date.now();
            while (Date.now() - startWait < 120000) {
                const isChatReady = await this.page.$("#contact-search-input, .conv-item, #main-tab").catch(() => null);
                if (isChatReady) {
                    this.isLoggedIn = true;
                    this.currentQrBase64 = null;
                    this.progress.qrImage = null;
                    this.progress.status = "opening_group";
                    this._log("Đăng nhập Zalo Web thành công! Đã lưu phiên làm việc.");
                    await new Promise(r => setTimeout(r, 2000));
                    return true;
                }
                await new Promise(r => setTimeout(r, 2000));
            }

            throw new Error("Hết thời gian chờ quét mã QR đăng nhập (2 phút). Vui lòng bấm Tải Lại và quét mã.");
        }

        // Nếu đã có giao diện chat
        const isChatReady = await this.page.$("#contact-search-input, .conv-item, #main-tab").catch(() => null);
        if (isChatReady) {
            this.isLoggedIn = true;
            this.currentQrBase64 = null;
            this.progress.qrImage = null;
            this._log("Đã kết nối tài khoản Zalo Web thành công.");
            return true;
        }

        throw new Error("Không thể kết nối vào Zalo Web. Vui lòng bật 'Hiện cửa sổ trình duyệt' và thử lại, hoặc kiểm tra mạng.");
    }

    /**
     * Main flow to download group photos
     */
    async downloadGroupPhotos(options = {}) {
        const {
            groupId,
            groupName,
            projectName,
            fromDate,
            toDate,
            folderFormat = "DD",
            headless = true,
            maxImages = 200
        } = options;

        if (this.isDownloading) {
            throw new Error("Tiến trình tải ảnh khác đang chạy, vui lòng chờ.");
        }

        if ((fromDate && !validDate(fromDate)) || (toDate && !validDate(toDate)) || (fromDate && toDate && fromDate > toDate)) {
            throw new Error("Khoảng ngày tải ảnh không hợp lệ.");
        }
        this.isDownloading = true;
        const targetProj = projectName || groupName || "Dự án mới";
        this.ensureProjectStructure(targetProj);

        this.progress = {
            status: "opening_group",
            current: 0,
            total: 0,
            downloaded: 0,
            skipped: 0,
            failed: 0,
            lowQuality: 0,
            unknownDate: 0,
            filteredOut: 0,
            counterVersion: 2,
            log: [],
            error: null,
            targetFolder: path.join(this.inputImagesDir, targetProj)
        };

        this._log(`Bắt đầu tải ảnh cho nhóm: "${groupName}" -> Dự án: "${targetProj}"`);
        if (fromDate && toDate) {
            this._log(`Lọc thời gian: Từ ${fromDate} đến ${toDate}`);
        }

        try {
            // Step 1: Ensure logged in
            await this.ensureLoggedIn(headless);

            // Step 2: Search and open group
            this.progress.status = "opening_group";
            await this._openGroup(groupName);

            // Step 3: Open Media Store (Ảnh/Video tab)
            this.progress.status = "scanning_media";
            await this._openMediaStore();

            // Resolve/download each batch while its virtualized tiles still exist.
            await this._collectPhotos(fromDate, toDate, maxImages, async photos => {
                this.progress.status = "downloading";
                await this._downloadPhotos(photos, targetProj, folderFormat, fromDate, toDate);
                this.progress.status = "scanning_media";
            });

            this.progress.status = "done";
            this._log(`Hoàn thành tải ảnh! Tải về: ${this.progress.downloaded}, Bỏ qua (đã có): ${this.progress.skipped}, Lỗi: ${this.progress.failed}`);
            return this.progress;

        } catch (err) {
            this._log(`[Lỗi] ${err.message}`);
            this.progress.status = "error";
            this.progress.error = err.message;
            throw err;
        } finally {
            this.isDownloading = false;
        }
    }

    /**
     * Search and click on the group in the conversation list
     */
    async _openGroup(groupName) {
        this._log(`Đang tìm kiếm và mở nhóm "${groupName}"...`);
        const page = this.page;

        // Check if group is already currently open
        const isCurrentChat = await page.evaluate((gName) => {
            const header = document.querySelector(".header-title, .chat-title, .title-text, .header-name");
            if (header && header.textContent && header.textContent.toLowerCase().includes(gName.toLowerCase())) {
                return true;
            }
            return false;
        }, groupName);

        if (isCurrentChat) {
            this._log(`Đang ở sẵn trong cuộc trò chuyện nhóm "${groupName}".`);
            return;
        }

        // Try direct click if already visible in recent list
        const clicked = await page.evaluate((gName) => {
            const items = document.querySelectorAll(".conv-item, .conv-item-title__name, .truncate");
            for (const el of items) {
                if (el.textContent && el.textContent.trim().toLowerCase() === gName.trim().toLowerCase()) {
                    el.closest(".conv-item")?.click() || el.click();
                    return true;
                }
            }
            return false;
        }, groupName);

        if (clicked) {
            this._log(`Đã chọn nhóm "${groupName}" từ danh sách gần đây.`);
            await new Promise(r => setTimeout(r, 2000));
            return;
        }

        // If not in view, use the search box
        const searchInput = await page.$("#contact-search-input, input[placeholder*='Tìm kiếm']");
        if (searchInput) {
            await searchInput.click();
            await page.keyboard.down("Control");
            await page.keyboard.press("A");
            await page.keyboard.up("Control");
            await page.keyboard.press("Backspace");

            await searchInput.type(groupName, { delay: 50 });
            this._log(`Đã gõ từ khóa tìm kiếm: "${groupName}"`);
            await new Promise(r => setTimeout(r, 2000));

            // Click the first matching result
            const resultClicked = await page.evaluate((gName) => {
                const results = document.querySelectorAll(".conv-item, .search-item, [data-id*='conv_item']");
                for (const item of results) {
                    if (item.textContent && item.textContent.toLowerCase().includes(gName.toLowerCase())) {
                        item.click();
                        return true;
                    }
                }
                if (results.length > 0) {
                    results[0].click();
                    return true;
                }
                return false;
            }, groupName);

            if (resultClicked) {
                this._log(`Đã mở cuộc trò chuyện nhóm "${groupName}".`);
                await new Promise(r => setTimeout(r, 2500));
                return;
            }
        }

        throw new Error(`Không tìm thấy nhóm "${groupName}" trên Zalo Web. Vui lòng đảm bảo tài khoản đã tham gia nhóm này.`);
    }

    /**
     * Open the right sidebar and select "Ảnh/Video" tab, then click "Xem tất cả" to open Kho lưu trữ
     */
    async _openMediaStore() {
        this._log("Đang mở Kho Media (Ảnh/Video) của nhóm...");
        const page = this.page;

        // Ensure right sidebar is open
        await page.evaluate(() => {
            const btn = document.querySelector("div[title='Thông tin hội thoại'], div[title='Thông tin nhóm']");
            if (btn && !btn.className.includes("focused")) {
                btn.click();
            }
        });
        await new Promise(r => setTimeout(r, 2000));

        // Click "Xem tất cả" inside the "Ảnh/Video" section of the sidebar
        const openedKho = await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll("*")).filter(el => {
                const txt = el.textContent ? el.textContent.trim() : "";
                return (txt === "Xem tất cả" || txt === "Xem tất cả >") && el.children.length === 0;
            });
            if (btns.length > 0) {
                btns[0].click();
                return true;
            }
            return false;
        });

        if (openedKho) {
            this._log("Đã mở Kho lưu trữ Ảnh/Video đầy đủ của nhóm.");
        } else {
            this._log("Đang quét ảnh trực tiếp từ Kho lưu trữ & Cuộc trò chuyện...");
        }

        await new Promise(r => setTimeout(r, 2500));
    }

    /**
     * Collect photos by scrolling the media store and chat view
     */
    async _collectPhotos(fromDateStr, toDateStr, maxPhotos = 200, onBatch = null) {
        this._log("Đang quét danh sách ảnh trong nhóm...");
        const limit = Math.max(1, Number(maxPhotos) || 200);
        const collected = [];
        const seenIds = new Set();
        const seenUrls = new Map();
        for (let scroll = 0; scroll < 8 && collected.length < limit; scroll++) {
            const raw = await this.page.evaluate(collectVisiblePhotos);
            const batch = mergePhotos(raw.map(photo => {
                const date = normalizeSendDate(photo);
                return {...photo, date, dateSource: photo.timestamp && date ? "message" : photo.dateSource};
            })).filter(photo => {
                if (photo.date && ((fromDateStr && photo.date < fromDateStr) || (toDateStr && photo.date > toDateStr))) return false;
                if (seenIds.has(photo.messageId ? photo.id : `${photo.id}:${photo.date}`)) return false;
                const previous = seenUrls.get(photo.url) || [];
                return !previous.some(p => p.date === photo.date && (!p.messageId || !photo.messageId));
            }).slice(0, limit - collected.length);
            for (const photo of batch) {
                seenIds.add(photo.messageId ? photo.id : `${photo.id}:${photo.date}`);
                seenUrls.set(photo.url, [...(seenUrls.get(photo.url) || []), photo]);
            }
            collected.push(...batch);
            this.progress.total = collected.length;
            if (batch.length && onBatch) await onBatch(batch);
            this._log(`Lần quét #${scroll + 1}: Tìm thấy ${collected.length} ảnh...`);
            if (collected.length >= limit) break;
            await this.page.evaluate(() => {
                const mediaScroll = document.querySelector("#innerScrollContainer")?.parentElement ||
                    document.querySelector(".media-store-view, .chat-right-menu, .chat-right-menu-content");
                if (mediaScroll) mediaScroll.scrollTop += 800;
                const chatScroll = document.querySelector("#messageViewContainer, .chat-message-list, #chatViewContainer, .chat-content");
                if (chatScroll) chatScroll.scrollTop -= 600;
            });
            await new Promise(r => setTimeout(r, 1500));
        }
        return collected;
    }

    async _resolvePhoto(photo) {
        return resolvePhoto(this.page, photo);
    }

    async _fetchImage(url) {
        return readImage(this.page, {url});
    }

    async _downloadPhotos(photos, projectName, folderFormat = "DD", fromDate, toDate) {
        const safeProject = (projectName || "Chung").replace(/[<>:"/\\|?*]/g, "_").trim();
        if (!safeProject || safeProject === "." || safeProject === "..") throw new Error("Tên dự án không hợp lệ");
        for (const initial of photos) {
            this.progress.current++;
            this.progress.currentFile = "";
            let photo = initial;
            let date = normalizeSendDate(photo);
            if (date && ((fromDate && date < fromDate) || (toDate && date > toDate))) {
                this.progress.filteredOut++;
                continue;
            }
            try {
                photo = await this._resolvePhoto(photo);
                date = normalizeSendDate(photo);
                if (!date) {
                    this.progress.unknownDate++;
                    this._log("[Cảnh báo] Bỏ qua ảnh: không xác định được ngày gửi trên Zalo.");
                    continue;
                }
                if ((fromDate && date < fromDate) || (toDate && date > toDate)) {
                    this.progress.filteredOut++;
                    continue;
                }
                let image = photo.fullImage;
                let lowQuality = false;
                let reason = photo.qualityWarning || "Zalo chưa cung cấp bản đầy đủ có thể xác minh";
                if (!image && photo.downloadUrl) {
                    try { image = await this._fetchImage(photo.downloadUrl); }
                    catch (error) { reason = error.message; }
                }
                if (!image) {
                    lowQuality = true;
                    image = await this._fetchImage(photo.url);
                }
                const day = folderFormat === "day" || folderFormat === "DD" ? date.slice(-2) : date;
                const directory = path.join(this.inputImagesDir, safeProject, day);
                fs.mkdirSync(directory, {recursive: true});
                this.progress.targetFolder = directory;
                // Start after the highest existing number, not the file count; never overwrite.
                let index = fs.readdirSync(directory).reduce((max, name) => {
                    const match = name.match(/^zalo_.+_(\d+)\.[^.]+$/);
                    return match ? Math.max(max, Number(match[1])) : max;
                }, 0);
                let filename, filePath;
                while (true) {
                    filename = `zalo_${day}_${String(++index).padStart(3, "0")}.${image.extension}`;
                    filePath = path.join(directory, filename);
                    try {
                        const fd = fs.openSync(filePath, "wx");
                        try { fs.writeFileSync(fd, image.buffer); }
                        catch (error) {
                            fs.closeSync(fd);
                            fs.unlinkSync(filePath);
                            throw error;
                        }
                        fs.closeSync(fd);
                        break;
                    } catch (error) {
                        if (error.code !== "EEXIST") throw error;
                    }
                }
                this.progress.currentFile = filename;
                this.progress.downloaded++;
                if (lowQuality) {
                    this.progress.lowQuality++;
                    this._log(`[Cảnh báo] ${filename}: chỉ lưu ảnh xem trước; ${reason}.`);
                }
                this._log(`[✓] ${filename}: ${image.width} × ${image.height} px, ${(image.buffer.length / 1024).toFixed(1)} KB; ngày gửi ${date} (giờ Việt Nam).`);
            } catch (error) {
                this.progress.failed++;
                this._log(`[✗] Không tải được ảnh: ${error.message}`);
            }
        }
    }

    async close() {
        if (this.browser) {
            try {
                await this.browser.close();
            } catch (e) {}
            this.browser = null;
            this.page = null;
        }
    }
}

module.exports = ZaloBrowserDownloader;
