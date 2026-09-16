/**
 * Image Downloader
 * Downloads images from Zalo and saves to input_images/{groupName}/{date}/
 * with deduplication support
 */

const fs = require("fs");
const path = require("path");
const https = require("https");
const http = require("http");
const {normalizeSendDate} = require("./photo-metadata");

const BASE_DIR = path.resolve(__dirname, "..");
const INPUT_IMAGES_DIR = path.join(BASE_DIR, "input_images");
const DOWNLOAD_LOG_PATH = path.join(__dirname, "download_log.json");

class ImageDownloader {
    constructor() {
        this.downloadLog = this._loadDownloadLog();
        this.progress = {
            total: 0,
            downloaded: 0,
            skipped: 0,
            failed: 0,
            status: "idle", // idle | downloading | done | error
            currentFile: "",
            errors: [],
        };
        this.progressListeners = [];
    }

    /**
     * Load download log from file
     */
    _loadDownloadLog() {
        try {
            if (fs.existsSync(DOWNLOAD_LOG_PATH)) {
                return JSON.parse(fs.readFileSync(DOWNLOAD_LOG_PATH, "utf-8"));
            }
        } catch {}
        return {};
    }

    /**
     * Save download log to file
     */
    _saveDownloadLog() {
        try {
            fs.writeFileSync(
                DOWNLOAD_LOG_PATH,
                JSON.stringify(this.downloadLog, null, 2),
                "utf-8"
            );
        } catch (err) {
            console.error("[ImageDownloader] Failed to save download log:", err.message);
        }
    }

    /**
     * Check if an image has already been downloaded
     */
    _isAlreadyDownloaded(msgId, timestamp) {
        const key = `${msgId}_${timestamp}`;
        return !!this.downloadLog[key];
    }

    /**
     * Mark image as downloaded
     */
    _markAsDownloaded(msgId, timestamp, filePath) {
        const key = `${msgId}_${timestamp}`;
        this.downloadLog[key] = {
            filePath,
            downloadedAt: new Date().toISOString(),
        };
    }

    /**
     * Add a progress event listener
     */
    onProgress(callback) {
        this.progressListeners.push(callback);
    }

    /**
     * Emit progress update
     */
    _emitProgress() {
        const data = { ...this.progress };
        for (const listener of this.progressListeners) {
            try { listener(data); } catch {}
        }
    }

    /**
     * Format date using local timezone
     */
    _formatDate(timestamp) {
        const ymd = normalizeSendDate({timestamp});
        return {ymd, day: ymd.slice(-2), month: ymd.slice(5, 7), year: ymd.slice(0, 4)};
    }

    /**
     * Download images from a group for a specific date range
     * @param {Array} images - array of { msgId, url, thumb, timestamp, sender, senderName }
     * @param {string} groupName - group name for folder
     * @param {string|null} dateFrom - YYYY-MM-DD format, null for no filter
     * @param {string|null} dateTo - YYYY-MM-DD format, null for no filter
     * @param {string} folderFormat - 'date' or 'YYYY-MM-DD' (full calendar date)
     */
    async downloadImages(images, groupName, dateFrom = null, dateTo = null, folderFormat = "date") {
        if (!["date", "YYYY-MM-DD"].includes(folderFormat)) throw new Error("Chỉ hỗ trợ YYYY-MM-DD");
        // Sanitize group name for folder
        const safeGroupName = this._sanitizeFolderName(groupName);

        // Filter by date range
        let filteredImages = images.filter(img => !!this._formatDate(img.timestamp).ymd);
        if (dateFrom || dateTo) {
            filteredImages = filteredImages.filter((img) => {
                const imgDateStr = this._formatDate(img.timestamp).ymd;
                if (dateFrom && imgDateStr < dateFrom) return false;
                if (dateTo && imgDateStr > dateTo) return false;
                return true;
            });
        }

        // Reset progress
        this.progress = {
            total: filteredImages.length,
            downloaded: 0,
            skipped: 0,
            failed: 0,
            status: "downloading",
            currentFile: "",
            errors: [],
        };
        this._emitProgress();

        console.log(`[ImageDownloader] Starting download: ${filteredImages.length} images for group "${groupName}"`);

        for (const img of filteredImages) {
            // Check dedup
            if (this._isAlreadyDownloaded(img.msgId, img.timestamp)) {
                this.progress.skipped++;
                this.progress.downloaded++;
                this._emitProgress();
                continue;
            }

            // Determine date folder
            const dateInfo = this._formatDate(img.timestamp);
            const dateFolder = dateInfo.ymd;

            // Create directory structure
            const targetDir = path.join(INPUT_IMAGES_DIR, safeGroupName, dateFolder);
            fs.mkdirSync(targetDir, { recursive: true });

            // Generate filename
            const ts = img.timestamp.toString();
            const msgId = (img.msgId || "unknown").toString().replace(/[^a-zA-Z0-9]/g, "");
            const filename = `${ts}_${msgId}.jpg`;
            const filePath = path.join(targetDir, filename);

            this.progress.currentFile = filename;
            this._emitProgress();

            try {
                // Try HD URL first, then thumb
                const url = img.url || img.thumb;
                if (!url) {
                    this.progress.failed++;
                    this.progress.errors.push(`No URL for message ${img.msgId}`);
                    this._emitProgress();
                    continue;
                }

                await this._downloadFile(url, filePath);
                fs.writeFileSync(filePath + ".json", JSON.stringify({date_source: "message",
                    send_date: dateFolder, message_id: img.msgId || "", derived: false}), "utf8");
                this._markAsDownloaded(img.msgId, img.timestamp, filePath);
                this.progress.downloaded++;
                console.log(`[ImageDownloader] Downloaded: ${filename}`);
            } catch (err) {
                this.progress.failed++;
                this.progress.errors.push(`Failed to download ${filename}: ${err.message}`);
                console.error(`[ImageDownloader] Failed: ${filename} - ${err.message}`);
            }

            this._emitProgress();
        }

        // Save download log
        this._saveDownloadLog();

        this.progress.status = "done";
        this.progress.currentFile = "";
        this._emitProgress();

        console.log(
            `[ImageDownloader] Done! Downloaded: ${this.progress.downloaded - this.progress.skipped}, ` +
            `Skipped: ${this.progress.skipped}, Failed: ${this.progress.failed}`
        );

        return this.progress;
    }

    /**
     * Download a single file from URL
     */
    _downloadFile(url, filePath) {
        return new Promise((resolve, reject) => {
            const protocol = url.startsWith("https") ? https : http;
            const timeout = 30000; // 30 seconds

            const request = protocol.get(url, { timeout }, (response) => {
                // Handle redirects
                if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                    this._downloadFile(response.headers.location, filePath)
                        .then(resolve)
                        .catch(reject);
                    return;
                }

                if (response.statusCode !== 200) {
                    reject(new Error(`HTTP ${response.statusCode}`));
                    return;
                }

                const file = fs.createWriteStream(filePath);
                response.pipe(file);

                file.on("finish", () => {
                    file.close();
                    // Verify file size
                    const stats = fs.statSync(filePath);
                    if (stats.size < 1000) {
                        // Too small, probably an error
                        fs.unlinkSync(filePath);
                        reject(new Error("File too small, may not be a valid image"));
                    } else {
                        resolve();
                    }
                });

                file.on("error", (err) => {
                    fs.unlink(filePath, () => {});
                    reject(err);
                });
            });

            request.on("timeout", () => {
                request.destroy();
                reject(new Error("Request timeout"));
            });

            request.on("error", (err) => {
                fs.unlink(filePath, () => {});
                reject(err);
            });
        });
    }

    /**
     * Sanitize folder name for filesystem
     */
    _sanitizeFolderName(name) {
        return name
            .replace(/[<>:"/\\|?*]/g, "_")
            .replace(/\s+/g, " ")
            .trim();
    }

    /**
     * Get current download progress
     */
    getProgress() {
        return { ...this.progress };
    }
}

module.exports = ImageDownloader;
