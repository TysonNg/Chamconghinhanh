/**
 * Lightweight, zero-dependency EXIF Date Extractor
 * Reads DateTimeOriginal (0x9003), CreateDate (0x9004), or DateTime (0x0132)
 * from JPEG or TIFF image buffers.
 */

function extractExifDate(buffer) {
    if (!buffer || !Buffer.isBuffer(buffer) || buffer.length < 32) return null;

    // Check for JPEG SOI marker (0xFF, 0xD8)
    if (buffer[0] === 0xFF && buffer[1] === 0xD8) {
        let offset = 2;
        while (offset < buffer.length - 4) {
            if (buffer[offset] !== 0xFF) break;
            const marker = buffer[offset + 1];
            // APP1 marker (0xE1)
            if (marker === 0xE1) {
                const length = buffer.readUInt16BE(offset + 2);
                const app1Data = buffer.slice(offset + 4, offset + 2 + length);
                const date = parseApp1Exif(app1Data);
                if (date) return date;
                offset += 2 + length;
            } else if (marker === 0xDA || marker === 0xD9) {
                // SOS (Start of Scan) or EOI (End of Image) - stop parsing headers
                break;
            } else {
                const length = buffer.readUInt16BE(offset + 2);
                offset += 2 + length;
            }
        }
    } else {
        // Try direct TIFF parsing
        return parseTiffHeader(buffer, 0);
    }

    return null;
}

function parseApp1Exif(app1) {
    if (app1.length < 14) return null;
    // Must start with "Exif\0\0"
    if (app1.toString('ascii', 0, 6) !== 'Exif\0\0') return null;
    return parseTiffHeader(app1, 6);
}

function parseTiffHeader(buffer, tiffStart) {
    if (buffer.length < tiffStart + 8) return null;
    const isLittleEndian = buffer[tiffStart] === 0x49 && buffer[tiffStart + 1] === 0x49; // "II"
    const isBigEndian = buffer[tiffStart] === 0x4D && buffer[tiffStart + 1] === 0x4D;    // "MM"
    if (!isLittleEndian && !isBigEndian) return null;

    const readU16 = (o) => isLittleEndian ? buffer.readUInt16LE(tiffStart + o) : buffer.readUInt16BE(tiffStart + o);
    const readU32 = (o) => isLittleEndian ? buffer.readUInt32LE(tiffStart + o) : buffer.readUInt32BE(tiffStart + o);

    if (readU16(2) !== 42) return null;

    const ifd0Offset = readU32(4);
    if (ifd0Offset < 8 || tiffStart + ifd0Offset >= buffer.length) return null;

    let exifSubIfdOffset = null;
    let fallbackDate = null;

    // Read IFD0
    const ifd0Count = readU16(ifd0Offset);
    for (let i = 0; i < ifd0Count && (ifd0Offset + 2 + (i + 1) * 12) <= buffer.length - tiffStart; i++) {
        const entryOffset = ifd0Offset + 2 + i * 12;
        const tag = readU16(entryOffset);
        if (tag === 0x8769) { // Exif IFD Pointer
            exifSubIfdOffset = readU32(entryOffset + 8);
        } else if (tag === 0x0132) { // DateTime
            fallbackDate = readExifString(buffer, tiffStart, readU32(entryOffset + 4), readU32(entryOffset + 8));
        }
    }

    // Read Exif SubIFD if present
    if (exifSubIfdOffset && (tiffStart + exifSubIfdOffset < buffer.length)) {
        const subCount = readU16(exifSubIfdOffset);
        for (let i = 0; i < subCount && (exifSubIfdOffset + 2 + (i + 1) * 12) <= buffer.length - tiffStart; i++) {
            const entryOffset = exifSubIfdOffset + 2 + i * 12;
            const tag = readU16(entryOffset);
            if (tag === 0x9003 || tag === 0x9004) { // DateTimeOriginal or CreateDate
                const dateStr = readExifString(buffer, tiffStart, readU32(entryOffset + 4), readU32(entryOffset + 8));
                if (dateStr) {
                    const parsed = parseExifDateString(dateStr);
                    if (parsed) return parsed;
                }
            }
        }
    }

    if (fallbackDate) {
        return parseExifDateString(fallbackDate);
    }

    return null;
}

function readExifString(buffer, tiffStart, count, valOrOffset) {
    if (count <= 4) {
        const buf = Buffer.alloc(4);
        buf.writeUInt32BE(valOrOffset);
        return buf.toString('ascii', 0, count).replace(/\0/g, '').trim();
    }
    const offset = tiffStart + valOrOffset;
    if (offset + count > buffer.length) return null;
    return buffer.toString('ascii', offset, offset + count).replace(/\0/g, '').trim();
}

function parseExifDateString(str) {
    if (!str || typeof str !== 'string') return null;
    // Format is "YYYY:MM:DD HH:MM:SS" or "YYYY-MM-DD HH:MM:SS"
    const match = str.match(/^(\d{4})[:\-](\d{2})[:\-](\d{2})/);
    if (!match) return null;
    const [_, y, m, d] = match;
    const year = Number(y);
    const month = Number(m);
    const day = Number(d);
    if (year >= 2000 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        return `${y}-${m}-${d}`;
    }
    return null;
}

module.exports = { extractExifDate, parseExifDateString };
