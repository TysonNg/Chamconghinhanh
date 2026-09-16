(function (root) {
    function fullDate(period, day) {
        if (!/^\d{4}-\d{2}$/.test(period || '') || !/^\d{1,2}$/.test(String(day))) {
            throw new Error('Chọn đầy đủ tháng, năm và ngày');
        }
        const iso = period + '-' + String(day).padStart(2, '0');
        const parsed = new Date(iso + 'T00:00:00Z');
        if (!Number.isFinite(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== iso) {
            throw new Error('Ngày không hợp lệ');
        }
        return iso;
    }
    const api = {fullDate};
    if (typeof module !== 'undefined' && module.exports) module.exports = api;
    else root.AttendanceDate = api;
})(typeof window !== 'undefined' ? window : globalThis);
