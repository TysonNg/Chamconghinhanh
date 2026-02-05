# Phần Mềm Nhận Diện Khuôn Mặt & Trích Xuất Ngày Tháng

Phần mềm giúp nhận diện khuôn mặt từ ảnh, trích xuất ngày tháng từ watermark, và so sánh với database ảnh chân dung.

## 📁 Cấu Trúc Thư Mục

```
phan mem quet mat/
├── input_images/          # Đặt ảnh cần quét vào đây
├── database/              # Database ảnh chân dung
│   ├── Chi_Nhanh_1/       
│   │   ├── Nguyen_Van_A/
│   │   │   └── portrait.jpg
│   │   └── Tran_Van_B/
│   │       └── portrait.jpg
│   └── Chi_Nhanh_2/
│       └── ...
├── results/               # Kết quả xuất ra (Excel)
├── src/                   # Mã nguồn Python
├── templates/             # HTML templates
├── static/                # CSS, JS
└── requirements.txt       # Danh sách thư viện
```

## 🛠️ Cài Đặt

### Bước 1: Cài đặt Python dependencies

```bash
cd "d:\Projects\phan mem quet mat"
pip install -r requirements.txt
```

### Bước 2: Cài đặt Tesseract OCR (để đọc ngày tháng từ ảnh)

1. Tải Tesseract từ: https://github.com/UB-Mannheim/tesseract/wiki
2. Cài đặt vào `C:\Program Files\Tesseract-OCR\`
3. Thêm ngôn ngữ Tiếng Việt khi cài

Sau khi cài xong, cập nhật `requirements.txt`:
```
pytesseract>=0.3.8
```

Và cài đặt:
```bash
pip install pytesseract
```

### Bước 3: Cài đặt Face Recognition (tùy chọn - để nhận diện khuôn mặt)

Yêu cầu:
- Visual Studio Build Tools (C++ build tools)
- CMake

```bash
pip install cmake
pip install dlib
pip install face_recognition
```

## 🚀 Chạy Phần Mềm

```bash
cd "d:\Projects\phan mem quet mat"
python src/app.py
```

Mở trình duyệt và truy cập: **http://localhost:5000**

## 📖 Hướng Dẫn Sử Dụng

### 1. Thiết Lập Database Ảnh Chân Dung

Có 2 cách:

**Cách 1: Qua giao diện web**
1. Vào tab **Database**
2. Nhấn **Thêm** để tạo chi nhánh mới
3. Upload ảnh chân dung cho từng nhân viên

**Cách 2: Thủ công**
1. Tạo thư mục chi nhánh trong `database/`, ví dụ: `database/Chi_Nhanh_HCM/`
2. Trong mỗi chi nhánh, tạo thư mục cho từng người: `database/Chi_Nhanh_HCM/Nguyen_Van_A/`
3. Đặt ảnh chân dung vào thư mục của người đó
4. Vào web, nhấn **Quét Lại Database**

### 2. Upload Ảnh Cần Quét

Có 2 cách:

**Cách 1: Qua giao diện web**
1. Vào tab **Quét Ảnh**
2. Kéo thả ảnh vào vùng upload hoặc click để chọn file

**Cách 2: Thủ công**
1. Copy ảnh vào thư mục `input_images/`

### 3. Bắt Đầu Quét

1. Vào tab **Quét Ảnh**
2. Nhấn nút **🚀 Bắt Đầu Quét**
3. Theo dõi tiến độ xử lý
4. Khi hoàn thành, file Excel sẽ được tạo trong thư mục `results/`

### 4. Xem & Tải Kết Quả

1. Vào tab **Kết Quả**
2. Xem danh sách báo cáo
3. Nhấn **Tải về** để download file Excel

## 📊 Định Dạng Kết Quả Excel

| STT | Tên File | Ngày Giờ | Địa Điểm | Chi Nhánh | Tên Người | Độ Tin Cậy (%) |
|-----|----------|----------|----------|-----------|-----------|----------------|
| 1 | image001.jpg | 24/12/2025 08:43:36 | Q.7, TP.HCM | Chi_Nhanh_1 | Nguyen_Van_A | 95.2 |

## ⚙️ Cấu Hình

Chỉnh sửa file `src/config.py`:

```python
# Ngưỡng nhận diện (0.0 - 1.0, nhỏ hơn = chính xác hơn)
FACE_RECOGNITION_TOLERANCE = 0.6

# Số thread xử lý song song
MAX_WORKERS = 4

# Port web server
FLASK_PORT = 5000
```

## ❓ Xử Lý Lỗi

### Lỗi "Tesseract không tìm thấy"
- Kiểm tra Tesseract đã được cài đặt chưa
- Sửa đường dẫn trong `src/config.py`:
```python
TESSERACT_CMD = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### Lỗi "face_recognition module not found"
- Cần cài đặt Visual Studio Build Tools trước
- Chạy lại: `pip install face_recognition`

### Không nhận diện được khuôn mặt
- Đảm bảo ảnh chân dung trong database rõ nét, chỉ có 1 khuôn mặt
- Thử giảm `FACE_RECOGNITION_TOLERANCE` xuống 0.5

## 📝 License

MIT License
"# Chamconghinhanh"  git init git add README.md git commit -m "first commit" git branch -M main git remote add origin https://github.com/TysonNg/Chamconghinhanh.git git push -u origin main
