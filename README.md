# Phần Mềm Chấm Công & Quét Mặt Đối Soát (v2.1.0)

Phần mềm tự động hóa đối soát bảng chấm công (Excel / PDF), nhận diện khuôn mặt nhân viên từ ảnh camera hiện trường (bằng AI DeepFace Facenet512), chèn ảnh minh chứng vào file Word chi tiết từng người và tự động xuất Báo cáo Giải trình Tổng hợp A4.

---

## 📁 Cấu Trúc Thư Mục

```
phan mem quet mat/
├── input_images/          # Ảnh camera lưu theo Dự án/Ngày (VD: input_images/Du_An_A/2026-09-18/)
├── database/              # Ảnh chân dung mẫu của nhân viên theo chi nhánh
├── excel_uploads/         # File Excel chấm công gốc tải lên
├── excel_extracted/       # Dữ liệu Excel đã trích xuất
├── excel_persons/         # Bảng chấm công đã tách theo từng nhân viên
├── excel_face_output/     # Kết quả Word đã quét mặt & chèn ảnh camera (từ Excel)
├── pdf_uploads/           # File PDF chấm công gốc tải lên
├── pdf_extracted/         # File Word bảng chấm công đã tách từ PDF
├── pdf_face_output/       # Kết quả Word đã quét mặt & chèn ảnh camera (từ PDF)
├── supplement_data/       # Dữ liệu ảnh bổ sung và cấu hình watermark
├── src/                   # Mã nguồn Python backend
├── templates/             # Giao diện HTML
├── static/                # CSS, JavaScript
└── requirements.txt       # Danh sách thư viện Python
```

---

## 🛠️ Cài Đặt & Khởi Chạy

> 📌 **Hướng dẫn chi tiết từng bước cho máy mới**: Xem file [HD_CHAY_MAY_MOI.md](HD_CHAY_MAY_MOI.md)

### Yêu Cầu Tiên Quyết
- **Hệ điều hành**: Windows 10/11 (64-bit)
- **Python**: **Python 3.12 (64-bit)** *(khuyến nghị; không dùng Python 3.14+ vì chưa hỗ trợ thư viện AI)*
- **Node.js LTS**: *(tùy chọn; dùng cho module tải ảnh từ Zalo)*

### 🚀 Cách 1: Khởi chạy 1-Click (Khuyên dùng)
Nhấp đúp vào file:
```cmd
PhanMemQuetMat.bat
```
Script sẽ tự động:
1. Tạo môi trường ảo `.venv` độc lập.
2. Cài đặt toàn bộ thư viện cần thiết từ `requirements.txt`.
3. Cài đặt module Zalo Service (nếu máy có Node.js).
4. Khởi động server và tự động mở trình duyệt tại `http://127.0.0.1:5000`.

### 💻 Cách 2: Khởi chạy thủ công qua dòng lệnh
```bash
# 1. Tạo và kích hoạt môi trường ảo
py -3.12 -m venv .venv
.venv\Scripts\activate

# 2. Cài đặt thư viện Python
pip install --upgrade pip
pip install -r requirements.txt

# 3. Cài đặt Zalo Service (nếu dùng tính năng Zalo)
cd zalo-service
npm install
cd ..

# 4. Khởi chạy ứng dụng
python main.py
```
Mở trình duyệt và truy cập: **http://127.0.0.1:5000**

---

## 📖 Hướng Dẫn Sử Dụng

### 1. Quản Lý Ảnh
- **Ảnh camera theo ngày:** Vào tab **Quản Lý Ảnh** -> **Ảnh Camera Theo Ngày**. Chọn ngày trong tháng và tải ảnh camera hiện trường lên (hoặc dùng tính năng **Tải Ảnh Zalo** để tự động kéo ảnh về thư mục ngày).
- **Ảnh chân dung nhân viên:** Vào tab **Quản Lý Ảnh** -> **Ảnh Chân Dung Nhân Viên**. Thêm chi nhánh và upload ảnh chân dung rõ mặt của từng nhân viên để làm ảnh mẫu đối soát.

### 2. Xử Lý Chấm Công & Quét Mặt
- **Đối với bảng chấm công Excel:**
  1. Vào tab **Xử Lý Chấm Công** -> chọn tab con **Excel**.
  2. Tải file Excel chấm công lên và bấm **"Tách theo từng người"**.
  3. Tại bảng danh sách đợt đã tách, bấm nút màu xanh **`[Quét mặt]`**.
  4. Hệ thống sẽ tự động đối soát các ngày vắng/thiếu giờ/trùng giờ, dùng AI so khớp khuôn mặt từ ảnh camera và chèn ảnh minh chứng vào file Word cá nhân.
- **Đối với bảng chấm công PDF:**
  1. Vào tab **Xử Lý Chấm Công** -> chọn tab con **PDF**.
  2. Tải file PDF chấm công lên và bấm trích xuất.
  3. Tại danh sách đợt trích xuất, bấm nút **`[Quét mặt]`**.

### 3. Báo Cáo & Kết Quả
- Sau khi quét mặt xong, vào tab **Báo Cáo & Kết Quả**:
  - Xem và tải các file Word chi tiết của từng nhân viên (kèm ảnh camera đã chèn).
  - Tải file **Giải Trình Tổng Hợp** (khổ A4 ngang, gom tất cả nhân viên có ngày bất thường kèm hình ảnh thực tế và lý do giải trình).

### 4. Bổ Sung Ảnh (Nếu Cần)
- Vào tab **Bổ Sung Ảnh** để tạo ảnh bổ sung minh chứng cho các trường hợp thiếu ảnh hoặc cần điều chỉnh thông tin watermark ngày giờ.

---

## ⚙️ Cấu Hình Nâng Cao

Chỉnh sửa trong `src/config.py`:
- `FACE_RECOGNITION_TOLERANCE`: Ngưỡng nhận diện khuôn mặt (mặc định tối ưu hóa theo Facenet512).
- `FLASK_PORT`: Cổng máy chủ web nội bộ (mặc định: `5000`).

---

## 📝 License
MIT License
