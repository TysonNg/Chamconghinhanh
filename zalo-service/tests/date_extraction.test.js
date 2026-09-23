const test = require('node:test');
const assert = require('node:assert/strict');
const { extractExifDate, parseExifDateString } = require('../exif-extractor');
const { normalizeSendDate } = require('../photo-metadata');

test('parseExifDateString parses YYYY:MM:DD and YYYY-MM-DD correctly', () => {
    assert.equal(parseExifDateString('2026:08:15 10:30:00'), '2026-08-15');
    assert.equal(parseExifDateString('2026-08-20 14:15:22'), '2026-08-20');
    assert.equal(parseExifDateString('invalid'), null);
    assert.equal(parseExifDateString('1999:01:01 00:00:00'), null); // Before 2000
});

test('extractExifDate parses synthetic JPEG with APP1 EXIF buffer', () => {
    // Build a minimal valid JPEG with APP1 EXIF segment
    // SOI
    const soi = Buffer.from([0xFF, 0xD8]);
    
    // APP1 header
    // Marker 0xFFE1, length (2 bytes), "Exif\0\0"
    const exifHeader = Buffer.from('Exif\0\0', 'ascii');
    
    // TIFF header (Little Endian: "II", 42, offset to IFD0 = 8)
    const tiffHeader = Buffer.from([
        0x49, 0x49,             // "II"
        0x2A, 0x00,             // 42
        0x08, 0x00, 0x00, 0x00  // IFD0 offset: 8 bytes from TIFF start
    ]);

    // IFD0: 1 entry (Exif IFD pointer 0x8769 -> points to offset 26 from TIFF start)
    // entry count (2 bytes) = 1
    // tag (2 bytes) = 0x8769, type (2 bytes) = 4 (LONG), count (4 bytes) = 1, value (4 bytes) = 26
    const ifd0 = Buffer.alloc(2 + 12 + 4);
    ifd0.writeUInt16LE(1, 0); // 1 entry
    ifd0.writeUInt16LE(0x8769, 2); // Exif IFD tag
    ifd0.writeUInt16LE(4, 4);      // LONG
    ifd0.writeUInt32LE(1, 6);      // count 1
    ifd0.writeUInt32LE(26, 10);    // offset = 26
    ifd0.writeUInt32LE(0, 14);     // next IFD = 0

    // Exif SubIFD at offset 26:
    // 1 entry (DateTimeOriginal 0x9003 -> string at offset 44)
    const dateStr = '2026:08:15 08:30:00\0';
    const subIfd = Buffer.alloc(2 + 12 + 4 + dateStr.length);
    subIfd.writeUInt16LE(1, 0); // 1 entry
    subIfd.writeUInt16LE(0x9003, 2); // DateTimeOriginal
    subIfd.writeUInt16LE(2, 4);      // ASCII
    subIfd.writeUInt32LE(dateStr.length, 6); // count
    subIfd.writeUInt32LE(44, 10);    // offset = 44
    subIfd.writeUInt32LE(0, 14);     // next IFD = 0
    subIfd.write(dateStr, 18, 'ascii'); // at 26 + 18 = 44

    const tiffBody = Buffer.concat([tiffHeader, ifd0, subIfd]);
    const app1Data = Buffer.concat([exifHeader, tiffBody]);
    
    const app1Length = app1Data.length + 2;
    const app1Marker = Buffer.alloc(4);
    app1Marker[0] = 0xFF;
    app1Marker[1] = 0xE1;
    app1Marker.writeUInt16BE(app1Length, 2);

    const jpegBuffer = Buffer.concat([soi, app1Marker, app1Data, Buffer.from([0xFF, 0xD9])]);

    const result = extractExifDate(jpegBuffer);
    assert.equal(result, '2026-08-15');
});

test('normalizeSendDate correctly decodes timestamp from epoch milliseconds', () => {
    // 1790160181783 ms = 2026-09-23T10:43:01.783Z (Vietnam: 2026-09-23)
    const photo = {
        timestamp: 1790160181783,
        date: '',
        dateSource: 'message'
    };
    const date = normalizeSendDate(photo);
    assert.equal(date, '2026-09-23');
});
