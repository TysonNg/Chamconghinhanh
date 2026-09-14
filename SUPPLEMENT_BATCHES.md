# Hồ sơ ảnh bổ sung

Tab Bổ Sung Ảnh: nhập nhân viên/ngày hồ sơ, upload 1–31 ảnh, xem ảnh gốc,
sửa ngày hoặc ghi chú từng dòng, lưu, nhập người duyệt, duyệt và tải ZIP.
Ngày hồ sơ không xác nhận thời điểm chụp. Ảnh và EXIF giữ nguyên byte, không gọi AI.

SQLite nằm tại supplement_data/supplement_batches.sqlite3; ảnh nằm trong
supplement_data/supplement_output/<batch_id>. ZIP chứa ảnh gốc, manifest.json
(ngày hồ sơ, ghi chú, SHA-256, lịch sử), README giải thích ý nghĩa hồ sơ.
Không đưa ảnh vào input_images. Sửa thông tin sẽ yêu cầu duyệt lại.
SHA-256 được kiểm tra trước duyệt và xuất. Upload/process cũ trả HTTP 410;
không xóa file cũ, không tự nhập hồ sơ cũ thành ảnh gốc đã xác minh.

Giới hạn: 20MB/ảnh, 100MB/request, 40 triệu pixel/ảnh; JPEG/PNG/WebP/BMP.
Người duyệt tự khai, chưa xác thực tài khoản. Audit không chống người có quyền
sửa trực tiếp cả SQLite và ảnh. Cần xác thực, phân quyền và CSRF trước khi
mở cho mạng không tin cậy. Upload/ZIP đồng bộ, giữ dữ liệu trong RAM có giới hạn.
Chưa có migration hồ sơ cũ, job nền hay kiểm thử bản đóng gói PyInstaller.
Sao lưu SQLite và supplement_output khi ứng dụng đã dừng ghi.

Kiểm thử:
`python -m pytest tests/test_supplement_batches.py tests/test_supplement_persistence.py tests/test_supplement_ui.py -q`

`node --check static/js/supplement-batches.js`
