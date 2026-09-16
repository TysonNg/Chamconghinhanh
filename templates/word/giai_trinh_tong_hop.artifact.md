# Hợp đồng template giải trình tổng hợp

## Mục đích

Template `giai_trinh_tong_hop.docx` dùng để tạo một báo cáo giải trình chấm công tổng hợp theo dự án và khoảng ngày. Bố cục được tham chiếu từ file mẫu `GIẢI TRÌNH Chung cư TTĐ Tháng 07.2026.docx`, sau đó chuẩn hóa để in A4 dọc.

## Các trường thay thế

- `{{PROJECT_NAME}}`: tên dự án.
- `{{FROM_DATE}}`: ngày bắt đầu, định dạng `dd/mm/yyyy`.
- `{{TO_DATE}}`: ngày kết thúc, định dạng `dd/mm/yyyy`.
- `{{REPORT_DATE_LONG}}`: ngày lập báo cáo bằng tiếng Việt.
- `{{DATA_ROWS}}`: vị trí chèn các dòng bất thường; token này phải nằm trong dòng dữ liệu đầu tiên của bảng 5 cột.

## Ràng buộc bố cục

- Khổ A4 dọc, lề trái/phải 1,4 cm và lề trên/dưới 1,25 cm.
- Bảng dữ liệu có đúng 5 cột: Tên, Ngày, Giải trình, Hình ảnh thực tế, Ghi chú.
- Hàng tiêu đề phải lặp lại ở đầu mỗi trang.
- Mỗi hàng dữ liệu phải bật `cantSplit` để ảnh và nội dung không tách qua hai trang.
- Ảnh phải được co vừa khung tối đa 4,6 x 3,2 cm, giữ nguyên tỷ lệ.
- Khối kết luận và ký xác nhận phải được giữ cùng nhau trên trang cuối.
- Footer hiển thị `Trang X / Y`.

## Tạo lại template

Chạy `build_aggregate_report_template.py`. Script gọi cùng bộ dựng template mà ứng dụng sử dụng, vì vậy file DOCX và mã xuất báo cáo luôn đồng bộ.

## Kiểm tra bản in

Đã kiểm tra bằng Microsoft Word với báo cáo thử 14 dòng trên 3 trang: tiêu đề bảng lặp đúng, không cắt hàng/ảnh, ảnh ngang và dọc không méo, phần ký nằm trọn trên trang cuối.
