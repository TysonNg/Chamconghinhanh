/**
 * SyncManager - Quản lý tự động tải ảnh Zalo & Tải bù ngày thiếu
 * Hỗ trợ lập lịch theo khung giờ hằng ngày và thuật toán đếm lùi 10 ngày phát hiện ngày thiếu ảnh.
 */

const fs = require("fs");
const path = require("path");

const vietnamDate = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Ho_Chi_Minh",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
});

const vietnamTime = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Ho_Chi_Minh",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
});

const IMAGE_EXTENSIONS = new Set([".jpg", ".jpeg", ".png", ".webp", ".bmp"]);

function getVietnamNowParts() {
    const d = new Date();
    const dateParts = Object.fromEntries(vietnamDate.formatToParts(d).map(p => [p.type, p.value]));
    const timeParts = Object.fromEntries(vietnamTime.formatToParts(d).map(p => [p.type, p.value]));
    return {
        year: parseInt(dateParts.year, 10),
        month: parseInt(dateParts.month, 10),
        day: parseInt(dateParts.day, 10),
        hour: parseInt(timeParts.hour, 10),
        minute: parseInt(timeParts.minute, 10),
        dateStr: `${dateParts.year}-${dateParts.month}-${dateParts.day}`,
        timeStr: `${timeParts.hour}:${timeParts.minute}`
    };
}

function getVietnamDateOffset(daysOffset = 0) {
    const now = getVietnamNowParts();
    const utcDate = new Date(Date.UTC(now.year, now.month - 1, now.day - daysOffset));
    const y = utcDate.getUTCFullYear();
    const m = String(utcDate.getUTCMonth() + 1).padStart(2, "0");
    const d = String(utcDate.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
}

class SyncManager {
    constructor(projectRootDir, browserDownloader = null) {
        this.projectRootDir = projectRootDir || path.resolve(__dirname, "..");
        this.inputImagesDir = path.join(this.projectRootDir, "input_images");
        this.configPath = path.join(__dirname, "sync_config.json");
        this.browserDownloader = browserDownloader;
        this.schedulerTimer = null;
        this.isRunningSync = false;
        this.lastTriggeredDate = null; // Ngày gần nhất đã kích hoạt lịch hàng ngày (YYYY-MM-DD)
        this.startupChecked = false;

        this.config = this.loadConfig();
    }

    /**
     * Tải cấu hình từ sync_config.json
     */
    loadConfig() {
        const defaultConfig = {
            enabled: false,
            scheduleTime: "22:00", // Giờ chạy cố định hằng ngày (HH:mm)
            lookbackDays: 10,       // Số ngày đếm lùi để phát hiện thiếu ảnh
            autoBackfillOnStartup: true, // Tự động phát hiện và tải bù khi mở app
            lastRun: null,
            lastStatus: "idle",
            lastLog: [],
            mappings: [] // Mảng: [{ groupId, groupName, projectName, active: true }]
        };

        if (fs.existsSync(this.configPath)) {
            try {
                const data = JSON.parse(fs.readFileSync(this.configPath, "utf-8"));
                return { ...defaultConfig, ...data };
            } catch (e) {
                console.error("[SyncManager] Lỗi đọc file cấu hình sync_config.json:", e.message);
            }
        }
        return defaultConfig;
    }

    /**
     * Lưu cấu hình vào sync_config.json
     */
    saveConfig(newConfig = {}) {
        this.config = { ...this.config, ...newConfig };
        try {
            fs.writeFileSync(this.configPath, JSON.stringify(this.config, null, 2), "utf-8");
            return true;
        } catch (e) {
            console.error("[SyncManager] Lỗi lưu cấu hình sync_config.json:", e.message);
            return false;
        }
    }

    /**
     * Đếm số lượng ảnh hợp lệ trong thư mục
     */
    countImagesInDir(dirPath) {
        if (!fs.existsSync(dirPath)) return 0;
        try {
            const files = fs.readdirSync(dirPath);
            let count = 0;
            for (const file of files) {
                const ext = path.extname(file).toLowerCase();
                if (IMAGE_EXTENSIONS.has(ext)) {
                    count++;
                }
            }
            return count;
        } catch {
            return 0;
        }
    }

    /**
     * Thuật toán phát hiện ngày thiếu ảnh (Gap Detection) trong N ngày gần nhất
     * @param {string} projectName Tên dự án
     * @param {number} lookbackDays Số ngày đếm lùi (mặc định 10)
     */
    detectGaps(projectName, lookbackDays = 10) {
        const days = Math.max(1, Math.min(31, parseInt(lookbackDays, 10) || 10));
        const safeName = (projectName || "").replace(/[<>:"/\\|?*]/g, "_").trim();
        const projectDir = path.join(this.inputImagesDir, safeName);

        const now = getVietnamNowParts();
        const todayStr = now.dateStr;

        const dateList = [];
        const missingDates = [];

        // Duyệt từ 0 (Hôm nay) lùi về N-1 ngày trước
        for (let i = 0; i < days; i++) {
            const dateStr = getVietnamDateOffset(i); // YYYY-MM-DD
            const [y, m, d] = dateStr.split("-");
            const displayDate = `${d}/${m}`;
            const isToday = (dateStr === todayStr);

            // Kiểm tra thư mục chuẩn YYYY-MM-DD
            const fullDateDir = path.join(projectDir, dateStr);
            let photoCount = this.countImagesInDir(fullDateDir);

            // Nếu thư mục YYYY-MM-DD chưa có, kiểm tra thêm thư mục legacy DD
            if (photoCount === 0) {
                const legacyDayDir = path.join(projectDir, d);
                const legacyCount = this.countImagesInDir(legacyDayDir);
                if (legacyCount > 0) {
                    photoCount = legacyCount;
                }
            }

            let status = "ok";
            if (isToday) {
                status = photoCount > 0 ? "ok" : "today_pending";
            } else {
                if (photoCount === 0) {
                    status = "missing";
                    missingDates.push(dateStr);
                }
            }

            dateList.push({
                date: dateStr,
                displayDate: displayDate,
                photoCount: photoCount,
                isToday: isToday,
                status: status
            });
        }

        // Sắp xếp ngày từ cũ nhất đến mới nhất để hiển thị trực quan
        dateList.reverse();
        missingDates.sort(); // Ngày thiếu cũ nhất sẽ ở đầu mảng

        const earliestMissingDate = missingDates.length > 0 ? missingDates[0] : null;

        return {
            projectName: safeName,
            lookbackDays: days,
            dates: dateList,
            missingDates: missingDates,
            hasMissing: missingDates.length > 0,
            earliestMissingDate: earliestMissingDate,
            today: todayStr
        };
    }

    /**
     * Bắt đầu bộ lập lịch chạy ngầm
     */
    startScheduler(onLog = null) {
        if (this.schedulerTimer) {
            clearInterval(this.schedulerTimer);
        }

        console.log(`[SyncManager] Bộ lập lịch đã khởi động. Giờ hẹn: ${this.config.scheduleTime || '22:00'}, Tự động: ${this.config.enabled ? 'BẬT' : 'TẮT'}`);

        // Kiểm tra sau 10 giây khởi động nếu có bật autoBackfillOnStartup
        setTimeout(() => {
            if (this.config.enabled && this.config.autoBackfillOnStartup && !this.startupChecked) {
                this.startupChecked = true;
                this.checkStartupBackfill();
            }
        }, 10000);

        // Vòng lặp kiểm tra định kỳ mỗi 30 giây
        this.schedulerTimer = setInterval(() => {
            this.checkDailySchedule();
        }, 30000);
    }

    stopScheduler() {
        if (this.schedulerTimer) {
            clearInterval(this.schedulerTimer);
            this.schedulerTimer = null;
        }
    }

    /**
     * Kiểm tra khi vừa khởi động: nếu hôm qua hoặc các ngày trước bị thiếu ảnh -> tự động tải bù
     */
    async checkStartupBackfill() {
        if (this.isRunningSync) return;
        if (!this.config.mappings || this.config.mappings.length === 0) return;

        // Nếu trình duyệt chưa đăng nhập Zalo Web, không tự động chạy ngầm tránh treo chờ mã QR
        if (this.browserDownloader && !this.browserDownloader.isLoggedIn) {
            console.log("[SyncManager] Trình duyệt chưa đăng nhập Zalo Web, tạm hoãn tự động tải bù khi khởi động để chờ người dùng đăng nhập.");
            return;
        }

        console.log("[SyncManager] Đang kiểm tra ngày thiếu ảnh khi khởi động hệ thống...");
        let needBackfill = false;
        for (const m of this.config.mappings) {
            if (!m.active) continue;
            const gaps = this.detectGaps(m.projectName, this.config.lookbackDays);
            if (gaps.hasMissing) {
                console.log(`[SyncManager] Dự án "${m.projectName}" bị thiếu ${gaps.missingDates.length} ngày ảnh (từ ${gaps.earliestMissingDate}). Sẽ tự động tải bù.`);
                needBackfill = true;
                break;
            }
        }

        if (needBackfill) {
            await this.runSync({ backfillMissing: true, reason: "Tự động tải bù ngày thiếu khi khởi động" });
        }
    }

    /**
     * Kiểm tra giờ hẹn hằng ngày
     */
    async checkDailySchedule() {
        if (!this.config.enabled) return;
        if (this.isRunningSync) return;
        if (!this.config.scheduleTime) return;

        const now = getVietnamNowParts();
        const currentTime = now.timeStr; // HH:mm
        const currentDate = now.dateStr; // YYYY-MM-DD

        // So khớp giờ và đảm bảo hôm nay chưa chạy
        if (currentTime === this.config.scheduleTime && this.lastTriggeredDate !== currentDate) {
            console.log(`[SyncManager] Đã tới khung giờ hẹn (${currentTime}) ngày ${currentDate}. Bắt đầu tự động tải ảnh...`);
            this.lastTriggeredDate = currentDate;
            await this.runSync({ backfillMissing: true, reason: `Đến giờ hẹn định kỳ hằng ngày (${currentTime})` });
        }
    }

    /**
     * Thực thi quét & tải ảnh cho tất cả các mapping đang hoạt động
     */
    async runSync(options = {}) {
        if (this.isRunningSync) {
            console.warn("[SyncManager] Tiến trình đồng bộ đang chạy, bỏ qua yêu cầu mới.");
            return { success: false, error: "Đang có tiến trình đồng bộ đang chạy" };
        }

        if (!this.browserDownloader) {
            console.error("[SyncManager] browserDownloader chưa được gắn.");
            return { success: false, error: "browserDownloader chưa sẵn sàng" };
        }

        if (this.browserDownloader.isDownloading) {
            console.warn("[SyncManager] Trình duyệt đang tải ảnh dở, bỏ qua.");
            return { success: false, error: "Trình duyệt đang bận tải ảnh" };
        }

        const activeMappings = (this.config.mappings || []).filter(m => m.active);
        if (activeMappings.length === 0) {
            console.log("[SyncManager] Không có nhóm Zalo nào được kích hoạt trong danh sách tự động.");
            return { success: false, error: "Chưa có nhóm nào được chọn để tự động tải" };
        }

        this.isRunningSync = true;
        const now = getVietnamNowParts();
        const today = now.dateStr;
        const syncLogs = [];
        const results = [];

        const log = (msg) => {
            const time = new Date().toLocaleTimeString("vi-VN", { hour12: false });
            const line = `[${time}] ${msg}`;
            console.log(`[SyncManager] ${line}`);
            syncLogs.push(line);
        };

        log(`Bắt đầu đồng bộ tự động cho ${activeMappings.length} nhóm Zalo (Lý do: ${options.reason || "Kích hoạt thủ công"})...`);

        try {
            for (const mapping of activeMappings) {
                log(`Đang kiểm tra nhóm: "${mapping.groupName}" -> Dự án: "${mapping.projectName}"`);

                const gaps = this.detectGaps(mapping.projectName, this.config.lookbackDays);
                let fromDate = today;
                let toDate = today;

                if (options.backfillMissing !== false && gaps.hasMissing && gaps.earliestMissingDate) {
                    fromDate = gaps.earliestMissingDate;
                    log(`Phát hiện thiếu ${gaps.missingDates.length} ngày (${gaps.missingDates.join(", ")}). Đặt khoảng tải từ ${fromDate} đến ${toDate} để tải bù.`);
                } else {
                    log(`Dữ liệu các ngày trước đã đầy đủ. Quét ảnh hôm nay (${today}).`);
                }

                // Gọi browserDownloader để tải ảnh
                try {
                    const downloadResult = await this.browserDownloader.downloadGroupPhotos({
                        groupId: mapping.groupId,
                        groupName: mapping.groupName,
                        projectName: mapping.projectName,
                        fromDate: fromDate,
                        toDate: toDate,
                        folderFormat: "YYYY-MM-DD",
                        headless: true,
                        maxImages: 0
                    });

                    results.push({
                        groupName: mapping.groupName,
                        projectName: mapping.projectName,
                        fromDate,
                        toDate,
                        downloaded: downloadResult.downloaded || 0,
                        skipped: downloadResult.skipped || 0,
                        failed: downloadResult.failed || 0
                    });

                    log(`Hoàn thành nhóm "${mapping.groupName}": Tải mới ${downloadResult.downloaded || 0} ảnh, Bỏ qua ${downloadResult.skipped || 0} ảnh.`);
                } catch (err) {
                    log(`[Lỗi] Không thể tải nhóm "${mapping.groupName}": ${err.message}`);
                    results.push({
                        groupName: mapping.groupName,
                        projectName: mapping.projectName,
                        error: err.message
                    });
                }
            }

            this.config.lastRun = new Date().toISOString();
            this.config.lastStatus = "success";
            this.config.lastResults = results;
            this.saveConfig();

            log("Đã hoàn thành toàn bộ chu trình đồng bộ tự động.");
            return { success: true, results, logs: syncLogs };
        } catch (globalErr) {
            log(`[Lỗi toàn cục] ${globalErr.message}`);
            this.config.lastStatus = "error";
            this.config.lastError = globalErr.message;
            this.saveConfig();
            return { success: false, error: globalErr.message, logs: syncLogs };
        } finally {
            this.isRunningSync = false;
        }
    }
}

module.exports = {
    SyncManager,
    getVietnamNowParts,
    getVietnamDateOffset
};
