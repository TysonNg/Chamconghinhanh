# Chuyển đổi dữ liệu chấm công

Bản cập nhật không tự chuyển thư mục hoặc đoán danh tính khi khởi động.

## Xác nhận danh tính

Trong **Ảnh → Chân dung**, chọn đúng dự án rồi bấm **Xác nhận mã** cho nguồn chân dung cũ. Nhập mã chấm công chính xác (giữ số 0 đầu), ngày bắt đầu thuộc dự án và người xác nhận. Không xác nhận nếu chưa phân biệt được người trùng tên.

Thêm nhân viên mới cũng yêu cầu ba thông tin này. Chuyển dự án cần ngày hiệu lực và mã tại dự án đích; giữ ảnh và lịch sử nguồn. Ngày kết thúc dự án nguồn không bao gồm ngày chuyển. Đổi tên dự án chỉ đổi tên hiển thị, không đổi ID hay thư mục.

Lưu trữ nhân viên ẩn danh tính khỏi danh sách hoạt động trên mọi dự án, giữ lịch sử. Ngừng dùng ảnh có hiệu lực từ ngày thao tác; ảnh gốc vẫn được giữ.

## Ngày ảnh và thư mục cũ

Ảnh mới lưu theo `dự án/YYYY-MM-DD`. Chọn tháng trước khi xem, tải lên hoặc xóa. Ngày gửi Zalo dùng để phân thư mục, không thay thế bằng chứng ngày chụp.

Thư mục `DD` chỉ dùng khi có mapping đã xác nhận trong `attendance-period.json`. Dùng nút gán ngày thư mục cũ, nhập ngày đầy đủ và người xác nhận. Nếu ảnh trộn nhiều tháng, phân loại thủ công vào các thư mục ISO riêng; không gán toàn bộ vào một tháng.

## Xem trước và áp dụng

Dừng ứng dụng và tác vụ tải ảnh trước khi áp dụng; không chạy nhiều công cụ chuyển đổi đồng thời. Sao lưu dữ liệu. Từ thư mục mã nguồn:

```powershell
python -m tools.preview_attendance_migration "input_images/Ten du an" --identity-db "data/identity.sqlite3"
```

Thay đường dẫn bằng thư mục thực tế. Mặc định chỉ đọc và in danh sách nguồn/đích/hash, thư mục chưa mapping, danh tính chưa xác nhận. Sau khi đối chiếu, thêm `--apply` để áp dụng.

Công cụ sao lưu manifest và DB nếu truyền DB, copy dữ liệu, kiểm tra SHA-256 rồi ghi nhật ký và cập nhật manifest. Giữ nguyên nguồn, từ chối ghi đè nội dung khác. Chạy lại không tạo bản sao trùng. Nếu nguồn thay đổi sau chuyển đổi, hệ thống yêu cầu kiểm tra.

Nếu gián đoạn, kiểm tra lỗi và chạy lại; không xóa nguồn. Khi khôi phục, dừng ứng dụng và dùng một bộ DB/manifest/thư mục thống nhất từ snapshot. Không chỉ phục hồi DB rồi chạy mã cũ trên ảnh đã chuẩn hóa.

## Trạng thái cần kiểm tra

- Thiếu mã xác nhận hoặc chân dung: không tự chọn theo tên.
- Ngày gốc khác hồ sơ: không tự đưa ảnh vào báo cáo.
- Chưa xác định ngày gốc: ảnh khớp mặt chỉ là ứng viên.
- Các nguồn cùng mã/ngày có giờ khác nhau: giữ mâu thuẫn để kiểm tra.
- Ảnh bổ sung: chỉ copy ảnh gốc có ngày phù hợp; ngày đề nghị hoặc watermark đã chỉnh không thay thế ngày gốc.

EXIF hoặc chữ ngày là thông tin đối chiếu, không chứng minh ảnh xác thực. Model và ngưỡng nhận diện không đổi.

## Kiểm chứng

Kiểm thử dùng thư mục tạm, ảnh tổng hợp và embedding giả. Chưa chuyển đổi dữ liệu thật, chạy phiên Zalo thật hoặc đánh giá độ chính xác model. Việc đóng/mở ứng dụng và kiểm tra bản EXE cần thực hiện trong quy trình phát hành.


## Tạo mã nội bộ hàng loạt

Trong **Ảnh → Chân dung**, bấm **Tạo mã nội bộ hàng loạt**. Nút áp dụng cho mọi nhân viên đang hoạt động trên tất cả dự án, không chỉ những người đang hiện trong kết quả tìm kiếm.

Mỗi người chưa có mã nội bộ được cấp mã ngẫu nhiên dạng NV-XXXXXXXX. Mã được kiểm tra không trùng trong cơ sở dữ liệu; bấm lại giữ nguyên mã đã có. Nếu trùng tên nhưng là hai hồ sơ nhân viên riêng, mỗi hồ sơ nhận một mã riêng. Khi chuyển dự án bằng chức năng chuyển nhân viên, mã nội bộ vẫn giữ nguyên.

**Mã nội bộ** dùng để quản lý trong ứng dụng. **Mã Excel/PDF** vẫn là mã trong bảng chấm công để đối chiếu danh tính và ngày hiệu lực. Tạo mã nội bộ không thay mã Excel/PDF, không tự xác nhận chân dung và không gán ngày làm việc.

Báo cáo xuất ra không hiển thị mã nội bộ hoặc mã chấm công; mã vẫn được sử dụng bên trong để phân biệt nhân viên.
