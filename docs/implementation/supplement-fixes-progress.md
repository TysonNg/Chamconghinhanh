# Sửa luồng ảnh bổ sung cho tool cá nhân

## Hành vi đã triển khai

- Popup HTML dùng chung thay native confirm; Đồng ý/Hủy/Escape/đóng, dọn overlay/listener/khóa cuộn, trả focus. Chặn bấm lặp và phục hồi nút kể cả lỗi mạng. Xóa tất cả chỉ gửi ID đang hiển thị; không có nhánh xóa mọi hồ sơ khi thiếu ID.
- Tạo ảnh qua API xử lý và archive bắt buộc project_id/employee_id. Tra registry, dự án còn hoạt động và membership tại ngày bổ sung; kiểm tra lại lúc áp dụng. Giao diện có membership lịch sử, chọn lại danh tính và kiểm tra lại hồ sơ cũ.
- Không tự sinh giờ. Thiếu giờ là null; explicit target_time:null trả 400. Chuẩn hóa giờ hợp lệ, đọc DateTimeOriginal cả Exif IFD; OCR đọc ảnh gốc không nhận ngày đề nghị. Ngày ảnh/EXIF/OCR mâu thuẫn chặn áp dụng.
- Nhận diện mọi mặt trong ảnh gốc so với chân dung đã xác nhận; giữ model/metric/ngưỡng hiện có. Không mặt, không chân dung, không khớp và lỗi model có lý do riêng. Cache và kết quả kiểm tra gắn với nội dung nguồn/chân dung/cấu hình.
- Kết quả ngày, mặt, toàn vẹn, lưu trữ, watermark và EXIF riêng. 201 chỉ xác nhận lưu hồ sơ. Watermark cục bộ cũng phải vượt hậu kiểm OCR; ảnh dẫn xuất bị đổi/mất không còn báo xử lý hoàn tất.
- Hash nguồn/ảnh dẫn xuất riêng. Áp dụng chỉ dùng ảnh gốc đạt kiểm tra; giữ nguồn và applied_at ban đầu. Chống bản sao theo dự án/ngày/hash, khác nhân viên báo xung đột.
- Apply/delete dùng khóa ghi chung xuyên thread/process, journal SQLite và file stage ngoài thư mục quét. Batch lỗi không công bố ảnh nào; reader/direct-image chỉ đọc thao tác committed. Phục hồi chỉ xóa file có bằng chứng sở hữu, không xóa file ngoại lai. Journal cũ thiếu bằng chứng sở hữu giữ nguyên file/metadata và báo unresolved.
- GET không xóa DB khi mất ảnh dẫn xuất. Có tạo lại ảnh từ nguồn nguyên vẹn. ZIP kiểm file/hash toàn bộ rồi xuất bytes đã kiểm tra; không bỏ qua file lỗi.
- Không cần duyệt; endpoint approve cũ trả 410. Metadata chỉ cập nhật dữ liệu/lịch sử và vô hiệu kiểm tra liên quan. Trạng thái duyệt cũ không chi phối thao tác.
- Migration phiên bản 2 chạy lại an toàn; backup SQLite và ảnh trước nâng cấp trong supplement_data/backups. Hồ sơ cũ không tự ghép tên, không coi hash mới là bằng chứng tiếp nhận ban đầu. Xác nhận nguồn cũ ghi rõ legacy_baseline.

## Kiểm chứng

Kết quả nghiệm thu ngày 2026-10-04: **287 kiểm thử Python đạt**, **38 kiểm thử Node đạt**, gồm Chrome/Edge thật, không skip. Kiểm tra cú pháp các module Python và JavaScript cùng git diff --check đạt.

Lệnh nghiệm thu:

```powershell
py -3.12 -m pytest tests -q -p no:cacheprovider --disable-warnings
node --test tests/*.cjs
git diff --check
```

Kiểm thử bao gồm injection lỗi file/metadata/DB, intake/apply/delete dở dang, đồng thời và race với file không thuộc thao tác, membership/danh tính, giờ/ngày/EXIF/OCR, nhóm nhiều mặt, cache stale, ZIP/hash, migration và popup Chrome/Edge thật. Request mạng trong kiểm thử giao diện được giả lập; không coi đây là tái hiện chính máy khác của người dùng.

Chạy model thật cục bộ trên 3 bản sao chân dung với registry/cache tạm: ArcFace, RetinaFace, cosine, ngưỡng 0.37. Hai ảnh tự so khớp đạt, một ảnh không có chân dung đơn mặt hợp lệ. Chi tiết trong supplement-live-smoke.json. Đây chỉ là kiểm tra đường chạy model, không phải benchmark độ chính xác. Chưa có bộ ảnh camera/chân dung được người dùng xác nhận nhãn nên chưa đo nhận nhầm và bỏ sót.

OCR AI trực tiếp chưa chạy: kiểm duyệt tự động yêu cầu cho phép riêng trước khi gửi chân dung thật tới provider bên ngoài. Không gửi ảnh để vượt qua chặn này. Chưa mô phỏng cắt điện thật; kiểm thử chỉ xác minh fsync/journal và phục hồi lỗi/dừng thao tác giả lập.

## Sử dụng và dữ liệu cũ

Khởi động lại tool từ mã nguồn/launcher để nạp thay đổi. EXE đóng gói trước đó chưa được build lại.

Hồ sơ cũ thiếu ID cần chọn lại dự án/nhân viên; thiếu hash tiếp nhận cần xác nhận nguồn hiện tại, rồi kiểm tra lại. Dòng “đã xác nhận nguồn hiện tại” không khẳng định ảnh chưa từng bị sửa trước đó. Ảnh đã áp dụng và lịch sử giữ nguyên. Không xóa backup/journal khi đang xử lý sự cố. Journal cũ báo unresolved cần đối chiếu backup và file thủ công vì không thể chứng minh quyền sở hữu chỉ từ sidecar.
