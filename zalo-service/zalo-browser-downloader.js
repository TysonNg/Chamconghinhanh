const fs = require("fs");
const path = require("path");
const puppeteer = require("puppeteer-core");
const {validDate, normalizeSendDate, mergePhotos, collectVisiblePhotos} = require("./photo-metadata");
const {readImage, resolvePhoto} = require("./photo-viewer");
const {extractExifDate} = require("./exif-extractor");

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
        this.scanDelayMs = 2000;

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
     * Đồng bộ phiên đăng nhập từ ZaloClient (credentials.json) vào Puppeteer
     */
    async syncSessionFromCredentials(credentials) {
        if (!credentials || !credentials.cookie || !Array.isArray(credentials.cookie) || credentials.cookie.length === 0) {
            return false;
        }
        try {
            this._log("Đang đồng bộ phiên đăng nhập Zalo vào trình duyệt...");
            await this.getOrLaunchBrowser(true); // khởi chạy ngầm

            const puppeteerCookies = credentials.cookie.map(c => {
                let domain = c.domain || ".zalo.me";
                if (!domain.includes("zalo.me")) domain = ".zalo.me";
                return {
                    name: c.key || c.name,
                    value: String(c.value || ""),
                    domain: domain,
                    path: c.path || "/",
                    httpOnly: !!c.httpOnly,
                    secure: c.secure !== false
                };
            }).filter(c => c.name && c.value);

            if (puppeteerCookies.length > 0) {
                await this.page.setCookie(...puppeteerCookies);
                this._log(`Đã nạp ${puppeteerCookies.length} cookies vào trình duyệt.`);

                // Mở thử trang chat.zalo.me để kích hoạt session
                await this.page.goto("https://chat.zalo.me", { waitUntil: "domcontentloaded", timeout: 25000 });
                await new Promise(r => setTimeout(r, 2500));

                const isChatReady = await this.page.$("#contact-search-input, .conv-item, #main-tab").catch(() => null);
                if (isChatReady) {
                    this.isLoggedIn = true;
                    this.currentQrBase64 = null;
                    this.progress.qrImage = null;
                    this.progress.qrExpiresAt = null;
                    this._log("Đồng bộ phiên Zalo Web thành công! Trình duyệt đã sẵn sàng tải ảnh.");
                    return true;
                }
            }
        } catch (err) {
            this._log(`Cảnh báo đồng bộ phiên: ${err.message}`);
        }
        return false;
    }

    /**
     * Bật/Tắt hiển thị cửa sổ Chrome thật cho người dùng xem
     */
    async setBrowserVisible(visible = true) {
        try {
            if (this.currentHeadless === !visible && this.browser && this.browser.connected) {
                if (visible && this.page) {
                    await this.page.bringToFront();
                    const { exec } = require("child_process");
                    const psFocus = `powershell -NoProfile -Command "$wshell = New-Object -ComObject wscript.shell; $procs = Get-Process -Name chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -match 'Zalo' }; if ($procs) { $wshell.AppActivate($procs[0].Id) }"`;
                    exec(psFocus);
                }
                return true;
            }

            // Khởi động lại ở chế độ hiển thị mong muốn
            this._log(`Chuyển trình duyệt sang chế độ: ${visible ? "Hiển thị cửa sổ Chrome" : "Chạy ngầm"}`);
            const wasUrl = this.page ? this.page.url() : "https://chat.zalo.me";
            await this.getOrLaunchBrowser(!visible);
            if (wasUrl && !wasUrl.startsWith("about:")) {
                await this.page.goto(wasUrl, { waitUntil: "domcontentloaded", timeout: 30000 });
            }
            return true;
        } catch (err) {
            this._log(`Lỗi khi chuyển đổi hiển thị trình duyệt: ${err.message}`);
            return false;
        }
    }

    /**
     * Ensure user is logged in to Zalo Web
     */
    async ensureLoggedIn(headless = true) {
        await this.getOrLaunchBrowser(headless);

        const currentUrl = this.page.url();
        if (!currentUrl.includes("chat.zalo.me") && !currentUrl.includes("id.zalo.me")) {
            this._log("Đang mở trang https://chat.zalo.me...");
            await this.page.goto("https://chat.zalo.me", { waitUntil: "domcontentloaded", timeout: 30000 });
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

        // Kiểm tra nhanh xem đã ở trạng thái đăng nhập hay chưa (tối đa 8 giây)
        let state = null; // 'logged_in' | 'needs_login'
        const startCheck = Date.now();
        while (Date.now() - startCheck < 10000) {
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

            await new Promise(r => setTimeout(r, 600));
        }

        // Nếu đã ở trong giao diện chat
        if (state === "logged_in") {
            this.isLoggedIn = true;
            this.currentQrBase64 = null;
            this.progress.qrImage = null;
            this.progress.qrExpiresAt = null;
            this._log("Đã kết nối tài khoản Zalo Web thành công.");
            await this._checkAndClickSyncPrompt();
            return true;
        }

        // Nếu chưa đăng nhập, điều hướng trực tiếp đến trang đăng nhập QR để tránh chờ redirect
        this._log("Chưa có phiên Zalo Web, đang tải trang đăng nhập mã QR...");
        const targetLoginUrl = "https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F";
        if (!this.page.url().includes("id.zalo.me")) {
            await this.page.goto(targetLoginUrl, { waitUntil: "domcontentloaded", timeout: 20000 }).catch(() => {});
        }

        this.isLoggedIn = false;
        this.progress.status = "waiting_qr";
        this.abortLogin = false;

        if (!headless || this.currentHeadless === false) {
            try {
                await this.page.bringToFront();
                const { exec } = require("child_process");
                const psFocus = `powershell -NoProfile -Command "$wshell = New-Object -ComObject wscript.shell; $procs = Get-Process -Name chrome -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -match 'Zalo' }; if ($procs) { $wshell.AppActivate($procs[0].Id) }"`;
                exec(psFocus);
            } catch (e) {}
        }

        this._log("👉 Vui lòng mở app Zalo trên điện thoại, quét mã QR hiển thị ngay trên giao diện web để đăng nhập.");

        // Chụp mã QR lần đầu (chuẩn hình vuông sắc nét)
        await this._captureQrImage();
        this.progress.qrExpiresAt = Date.now() + 60000;

        // Chờ người dùng quét mã trên điện thoại (tối đa 120s)
        this._log("Đang chờ quét mã trên điện thoại (tự động làm mới khi hết hạn, hoặc bấm 'Hiện Cửa Sổ Chrome')...");
        const startWait = Date.now();
        let lastCaptureTime = Date.now();

        while (Date.now() - startWait < 120000) {
            if (this.abortLogin) {
                throw new Error("Đã hủy quá trình chờ quét mã QR.");
            }

            const currentUrl = this.page.url();
            const isOnChatDomain = currentUrl.includes("chat.zalo.me") && !currentUrl.includes("id.zalo.me");

            // Kiểm tra xem có đang yêu cầu xác minh bảo mật không (Chọn 3 người bạn gần đây)
            const needsVerification = await this.page.evaluate(() => {
                const text = (document.body ? document.body.innerText : "").toLowerCase();
                return text.includes("chọn 3 người") || text.includes("xác thực tài khoản") || text.includes("xác minh danh tính") || text.includes("liên lạc gần đây");
            }).catch(() => false);

            if (needsVerification && this.currentHeadless) {
                this._log("⚠️ Zalo yêu cầu xác minh danh tính (chọn 3 người bạn). Đang tự động mở cửa sổ Chrome lên màn hình để bạn chọn...");
                await this.setBrowserVisible(true);
            }

            if (isOnChatDomain) {
                await this._checkAndClickSyncPrompt();
                const isChatReady = await this.page.$("#contact-search-input, .conv-item, #main-tab, .chat-message, .zl-avatar").catch(() => null);
                if (isChatReady) {
                    this.isLoggedIn = true;
                    this.currentQrBase64 = null;
                    this.progress.qrImage = null;
                    this.progress.qrExpiresAt = null;
                    this.progress.status = "opening_group";
                    this._log("Đăng nhập Zalo Web thành công! Đã lưu phiên làm việc.");
                    await this._checkAndClickSyncPrompt();
                    await new Promise(r => setTimeout(r, 1500));
                    return true;
                }
            } else {
                const isChatReady = await this.page.$("#contact-search-input, .conv-item, #main-tab").catch(() => null);
                if (isChatReady) {
                    this.isLoggedIn = true;
                    this.currentQrBase64 = null;
                    this.progress.qrImage = null;
                    this.progress.qrExpiresAt = null;
                    this.progress.status = "opening_group";
                    this._log("Đăng nhập Zalo Web thành công! Đã lưu phiên làm việc.");
                    await this._checkAndClickSyncPrompt();
                    await new Promise(r => setTimeout(r, 1500));
                    return true;
                }
            }

            // Tự động kiểm tra mã QR hết hạn trên Zalo Web và click làm mới
            try {
                const isExpired = await this.page.evaluate(() => {
                    const selectors = [
                        ".qrcode-expired",
                        ".qr-expired",
                        "[class*='expired']",
                        "[class*='refresh-qr']"
                    ];
                    for (const sel of selectors) {
                        const el = document.querySelector(sel);
                        if (!el) continue;
                        const style = window.getComputedStyle(el);
                        if (style.display !== "none" && style.visibility !== "hidden" && style.opacity !== "0") return sel;
                    }
                    return false;
                }).catch(() => false);

                if (isExpired) {
                    this._log("Mã QR Zalo đã hết hạn, đang tự động lấy mã mới...");
                    await this.page.evaluate(() => {
                        const refreshSelectors = [
                            ".qrcode-expired", ".qr-expired",
                            "[class*='expired']", "[class*='refresh-qr']",
                            "[class*='retry']", "button[class*='refresh']"
                        ];
                        for (const sel of refreshSelectors) {
                            const el = document.querySelector(sel);
                            if (el && el.getClientRects().length) { el.click(); return; }
                        }
                    }).catch(() => {});
                    await new Promise(r => setTimeout(r, 2000));
                    await this._captureQrImage();
                    lastCaptureTime = Date.now();
                    this.progress.qrExpiresAt = Date.now() + 60000;
                }
            } catch (e) {}

            // Định kỳ chụp lại nếu chưa có ảnh hoặc sau 6 giây
            if (!this.progress.qrImage || (Date.now() - lastCaptureTime > 6000)) {
                await this._captureQrImage();
                lastCaptureTime = Date.now();
            }

            await new Promise(r => setTimeout(r, 1200));
        }

        throw new Error("Hết thời gian chờ quét mã QR đăng nhập (2 phút). Vui lòng bấm Lấy Mã QR Mới hoặc bấm 'Hiện Cửa Sổ Chrome'.");
    }

    /**
     * Tự động phát hiện và nhấn nút Đồng bộ tin nhắn trên Zalo Web nếu có xuất hiện
     * Quét cả popup, banner thông báo và dialog xác nhận đồng bộ
     */
    async _checkAndClickSyncPrompt() {
        if (!this.page) return;
        try {
            const clicked = await this.page.evaluate(() => {
                // 1. Tìm theo selectors cụ thể của Zalo Web cho popup/banner đồng bộ
                const specificSelectors = [
                    ".sync-msg-btn",
                    ".sync-banner button",
                    ".sync-popup button",
                    "[data-id*='sync'] button",
                    "[class*='sync-msg']",
                    "[class*='sync-banner']",
                    "[class*='sync_btn']",
                    "div[title*='Đồng bộ']",
                    "button[title*='Đồng bộ']"
                ];

                for (const sel of specificSelectors) {
                    const el = document.querySelector(sel);
                    if (el && el.getClientRects().length > 0) {
                        const style = window.getComputedStyle(el);
                        if (style.display !== "none" && style.visibility !== "hidden") {
                            el.click();
                            return sel;
                        }
                    }
                }

                // 2. Tìm theo text "đồng bộ ngay", "đồng bộ tin nhắn", "bắt đầu đồng bộ", "đồng bộ dữ liệu"
                const candidates = Array.from(document.querySelectorAll("button, .btn, div[role='button'], a, span, p"));
                for (const el of candidates) {
                    const text = (el.textContent || "").trim().toLowerCase();
                    if (
                        (text === "đồng bộ ngay" ||
                         text === "đồng bộ tin nhắn" ||
                         text.includes("đồng bộ ngay") ||
                         text.includes("đồng bộ tin nhắn") ||
                         text === "bắt đầu đồng bộ" ||
                         text === "tiếp tục đồng bộ" ||
                         text === "đồng bộ dữ liệu" ||
                         text === "bỏ qua" ||
                         text === "để sau" ||
                         text === "lúc khác" ||
                         text === "không phải bây giờ" ||
                         text.includes("sync now")) &&
                        el.children.length <= 2
                    ) {
                        const clickable = el.closest("button, .btn, div[role='button'], a") || el;
                        clickable.click();
                        return text;
                    }
                }
                return false;
            });

            if (clicked) {
                this._log(`Đã phát hiện và tự động nhấn nút đồng bộ: "${clicked}".`);
                await new Promise(r => setTimeout(r, 2000));

                // 3. Kiểm tra xem có popup phụ xác nhận không: "Xác nhận" / "Đồng ý" / "Tiếp tục"
                await this.page.evaluate(() => {
                    const confirmBtns = Array.from(document.querySelectorAll(".modal button, .dialog button, [role='dialog'] button, .popup button"));
                    for (const btn of confirmBtns) {
                        const t = (btn.textContent || "").trim().toLowerCase();
                        if (t === "xác nhận" || t === "đồng ý" || t === "tiếp tục" || t === "bắt đầu") {
                            btn.click();
                            return t;
                        }
                    }
                    return false;
                });
                await new Promise(r => setTimeout(r, 2000));
            }
        } catch {}
    }

    /**
     * Tự động làm mới danh sách tin nhắn & ảnh trong nhóm trước khi quét
     * Kích hoạt Zalo Web fetch dữ liệu mới nhất từ server, tránh bị lưu cache DOM cũ
     */
    async _refreshGroupMessages() {
        if (!this.page) return;
        this._log("Đang làm mới danh sách tin nhắn và ảnh mới nhất trong nhóm...");
        try {
            await this.page.evaluate(() => {
                // 1. Cuộn cửa sổ chat xuống cuối để nhận tin nhắn mới nhất
                const chatScroll = document.querySelector("#messageViewContainer, .chat-message-list, #chatViewContainer, .chat-content");
                if (chatScroll) {
                    chatScroll.scrollTop = chatScroll.scrollHeight;
                    chatScroll.dispatchEvent(new Event('scroll', { bubbles: true }));
                }

                // 2. Click nút "Tin nhắn mới" hoặc "Cuộn xuống cuối" nếu xuất hiện
                const newMsgBtn = document.querySelector(".chat-new-msg, [class*='new-msg'], [class*='scroll-bottom'], div[title*='Tin nhắn mới'], [data-id*='new_msg']");
                if (newMsgBtn) {
                    newMsgBtn.click();
                }

                // 3. Nếu Kho Media đang mở từ trước, đóng lại để chuẩn bị mở mới sạch sẽ
                const sidebarBtn = document.querySelector("div[title='Thông tin hội thoại'], div[title='Thông tin nhóm']");
                if (sidebarBtn && sidebarBtn.className.includes("focused")) {
                    sidebarBtn.click(); // Đóng lại
                }
            });

            // Chờ 1.5 giây để Zalo Web cập nhật trạng thái DOM sạch sẽ
            await new Promise(r => setTimeout(r, 1500));
            this._log("Đã làm mới dữ liệu nhóm thành công.");
        } catch (e) {
            this._log(`Cảnh báo làm mới nhóm: ${e.message}`);
        }
    }

    /**
     * Chụp mã QR sắc nét, chuẩn hình vuông từ trang Zalo Web
     * Đảm bảo ảnh luôn vuông vức, không chụp viền ngoài làm co nhỏ mã
     */
    async _captureQrImage() {
        if (!this.page) return null;
        try {
            // Đợi selector xuất hiện
            await this.page.waitForSelector(".qr-container svg, .qrcode canvas, .qrcode img, .qr-container", { timeout: 8000 });

            // Kiểm tra xem mã QR có đang bị che bởi overlay hết hạn không
            const isOverlayExpired = await this.page.evaluate(() => {
                const expEl = document.querySelector(".qrcode-expired, .qr-expired, [class*='expired']");
                if (expEl) {
                    const style = window.getComputedStyle(expEl);
                    if (style.display !== "none" && style.visibility !== "hidden" && style.opacity !== "0") {
                        return true;
                    }
                }
                return false;
            }).catch(() => false);

            if (isOverlayExpired) {
                // Không chụp ảnh khi đang báo hết hạn để tránh hiển thị ảnh co nhỏ/lỗi
                return null;
            }

            // Ưu tiên 1: svg bên trong .qr-container (Zalo Web chuẩn)
            const svgEl = await this.page.$(".qr-container svg");
            let targetEl = svgEl;

            // Ưu tiên 2: canvas hoặc img
            if (!targetEl) {
                targetEl = await this.page.$(".qrcode canvas, .qrcode img, canvas.qrcode");
            }

            if (targetEl) {
                const b64 = await targetEl.screenshot({ encoding: "base64" });
                const dataUri = "data:image/png;base64," + b64;
                this.currentQrBase64 = dataUri;
                this.progress.qrImage = dataUri;
                try {
                    fs.writeFileSync(path.join(__dirname, "qr.png"), Buffer.from(b64, "base64"));
                } catch {}
                return dataUri;
            }

            // Fallback: Nếu chỉ tìm thấy .qr-container, clip lấy vùng hình vuông ở trên (chứa mã QR)
            const qrCont = await this.page.$(".qr-container");
            if (qrCont) {
                const box = await qrCont.boundingBox();
                if (box) {
                    const side = Math.min(box.width, box.height, 240);
                    const b64 = await this.page.screenshot({
                        clip: {
                            x: box.x + (box.width - side) / 2,
                            y: box.y,
                            width: side,
                            height: side
                        },
                        encoding: "base64"
                    });
                    const dataUri = "data:image/png;base64," + b64;
                    this.currentQrBase64 = dataUri;
                    this.progress.qrImage = dataUri;
                    try {
                        fs.writeFileSync(path.join(__dirname, "qr.png"), Buffer.from(b64, "base64"));
                    } catch {}
                    return dataUri;
                }
            }
        } catch (e) {
            console.warn("[BrowserDownloader] _captureQrImage warning:", e.message);
        }
        return null;
    }

    /**
     * Yêu cầu làm mới mã QR ngay lập tức từ người dùng
     */
    async refreshQrCode() {
        if (!this.page || this.progress.status !== "waiting_qr") {
            throw new Error("Không có tiến trình chờ quét QR nào đang chạy.");
        }
        this._log("Đang làm mới mã QR Zalo theo yêu cầu...");
        try {
            const clicked = await this.page.evaluate(() => {
                const refreshSelectors = [
                    ".qrcode-expired", ".qr-expired",
                    "[class*='expired']", "[class*='refresh-qr']",
                    "[class*='retry']", "button[class*='refresh']"
                ];
                for (const sel of refreshSelectors) {
                    const el = document.querySelector(sel);
                    if (el && el.getClientRects().length) {
                        el.click();
                        return true;
                    }
                }
                return false;
            });
            if (!clicked) {
                await this.page.goto("https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F", { waitUntil: "domcontentloaded", timeout: 20000 });
            }
        } catch (e) {
            await this.page.goto("https://id.zalo.me/account?continue=https%3A%2F%2Fchat.zalo.me%2F", { waitUntil: "domcontentloaded", timeout: 20000 }).catch(() => {});
        }
        await new Promise(r => setTimeout(r, 2000));
        const newQr = await this._captureQrImage();
        this.progress.qrExpiresAt = Date.now() + 60000;
        this._log("Đã cập nhật mã QR mới thành công.");
        return newQr;
    }

    /**
     * Hủy tiến trình tải hoặc chờ quét QR
     */
    async cancelDownload() {
        this.abortLogin = true;
        this.isDownloading = false;
        this.progress.status = "idle";
        this.progress.qrImage = null;
        this.currentQrBase64 = null;
        this._log("Đã hủy tiến trình tải ảnh / chờ quét mã QR.");
        await this.close();
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
            folderFormat = "YYYY-MM-DD",
            headless = true,
            maxImages = 1000
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
            await this._checkAndClickSyncPrompt();

            // Tự động làm mới danh sách tin nhắn & ảnh mới nhất trong nhóm, tránh cache DOM cũ
            await this._refreshGroupMessages();

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
    async _collectPhotos(fromDateStr, toDateStr, maxPhotos = 1000, onBatch = null) {
        this._log("Đang quét danh sách ảnh trong nhóm...");
        const limit = Math.max(1, Number(maxPhotos) || 1000);
        const collected = [];
        const seenIds = new Set();
        const seenUrls = new Map();
        const maxScrolls = fromDateStr ? 120 : 8;
        let stagnantScans = 0;
        let previousVisibleKey = "";
        for (let scroll = 0; scroll < maxScrolls && collected.length < limit; scroll++) {
            const raw = await this.page.evaluate(collectVisiblePhotos);
            const normalized = mergePhotos(raw.map(photo => {
                const date = normalizeSendDate(photo);
                return {...photo, date, dateSource: photo.timestamp && date ? "message" : photo.dateSource};
            }));
            const visibleDates = normalized.map(photo => photo.date).filter(validDate).sort();
            const visibleKey = normalized.map(photo => `${photo.id}:${photo.date}`).sort().join("|");
            stagnantScans = visibleKey && visibleKey === previousVisibleKey ? stagnantScans + 1 : 0;
            previousVisibleKey = visibleKey;
            const batch = normalized.filter(photo => {
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
            if (fromDateStr && visibleDates.length && visibleDates[0] < fromDateStr) break;
            const maxStagnant = fromDateStr ? 6 : 3;
            if (stagnantScans >= maxStagnant) {
                const footerNotice = await this.page.evaluate(() => {
                    const el = document.querySelector("#footer, .tds-media-list__footer-wrapper, .tds-media-list__footer-content");
                    return el ? (el.textContent || "").trim() : "";
                });
                if (footerNotice) {
                    this._log(`[Giới hạn Zalo] ${footerNotice}`);
                } else {
                    this._log("Đã quét và cuộn đến cuối lịch sử ảnh có sẵn trên Zalo Web.");
                }
                break;
            }
            await this.page.evaluate(() => {
                const isc = document.querySelector("#innerScrollContainer");
                if (isc) {
                    let scrollEl = isc;
                    while (scrollEl && scrollEl !== document.body) {
                        const style = window.getComputedStyle(scrollEl);
                        if (style.overflowY === 'auto' || style.overflowY === 'scroll') {
                            break;
                        }
                        scrollEl = scrollEl.parentElement;
                    }
                    if (scrollEl && scrollEl !== document.body) {
                        scrollEl.scrollTop += 800;
                        scrollEl.dispatchEvent(new Event('scroll', { bubbles: true }));
                    }
                }
                const chatScroll = document.querySelector("#messageViewContainer, .chat-message-list, #chatViewContainer, .chat-content");
                if (chatScroll) {
                    chatScroll.scrollTop -= 600;
                    chatScroll.dispatchEvent(new Event('scroll', { bubbles: true }));
                }
            });
            try {
                const mediaHandle = await this.page.$("#innerScrollContainer, .media-store-view, .chat-right-menu");
                if (mediaHandle) {
                    const box = await mediaHandle.boundingBox();
                    if (box && box.width > 0 && box.height > 0) {
                        const targetX = Math.min(Math.max(box.x + box.width / 2, 10), 1200);
                        const targetY = Math.min(Math.max(box.y + 200, 50), 700);
                        await this.page.mouse.move(targetX, targetY);
                        await this.page.mouse.wheel({ deltaY: 800 });
                    }
                }
            } catch {}
            await new Promise(r => setTimeout(r, this.scanDelayMs));
        }
        return collected;
    }

    async _resolvePhoto(photo) {
        return resolvePhoto(this.page, photo);
    }

    async _fetchImage(url) {
        return readImage(this.page, {url});
    }

    async _downloadPhotos(photos, projectName, folderFormat = "YYYY-MM-DD", fromDate, toDate) {
        if (!["date", "YYYY-MM-DD"].includes(folderFormat)) throw new Error("Chỉ hỗ trợ YYYY-MM-DD");
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

                // Tầng 3: Trích xuất ngày từ EXIF của ảnh gốc nếu Zalo chưa xác định được ngày
                if (!date && image && image.buffer) {
                    const exifDate = extractExifDate(image.buffer);
                    if (exifDate) {
                        date = exifDate;
                        photo.date = exifDate;
                        photo.dateSource = "exif";
                        this._log(`[Nhận diện ngày] Đã đọc được ngày từ EXIF ảnh gốc: ${date}`);
                    }
                }

                if (!date) {
                    this.progress.unknownDate++;
                    // Thay vì bỏ qua hoàn toàn, lưu ảnh vào thư mục "unknown-date"
                    date = "unknown-date";
                    this._log("[Cảnh báo] Không xác định được ngày gửi, ảnh sẽ lưu vào thư mục unknown-date.");
                }
                // Chỉ lọc theo ngày nếu date là ngày cụ thể (không phải unknown-date)
                if (date !== "unknown-date" && ((fromDate && date < fromDate) || (toDate && date > toDate))) {
                    this.progress.filteredOut++;
                    continue;
                }
                const day = date;
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
                fs.writeFileSync(filePath + ".json", JSON.stringify({
                    date_source: "message", send_date: date, message_id: photo.messageId || photo.id || "",
                    derived: false
                }), "utf8");
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
