# 📖 HƯỚNG DẪN CÀI ĐẶT & CHẠY DỰ ÁN TỪ ĐẦU TRÊN MÁY TÍNH MỚI

Tài liệu này hướng dẫn chi tiết từng bước giúp bạn đưa mã nguồn dự án về bất kỳ máy tính Windows mới nào và khởi chạy thành công ngay từ lần đầu tiên.

---

## 1. ⚙️ YÊU CẦU PHẦN MỀM TIÊN QUYẾT

Trước khi chạy phần mềm, hãy đảm bảo máy tính mới đã cài đặt các công cụ sau:

### 1.1. Python 3.12 (Bắt buộc - 64-bit)
- **Lưu ý đặc biệt quan trọng**: Hệ thống sử dụng các thư viện AI chuyên sâu (`TensorFlow`, `DeepFace`, `PyTorch` / `EasyOCR`, `OpenCV`). Các thư viện này **chỉ tương thích ổn định nhất trên Python 3.10 - 3.12**.
- **TUYỆT ĐỐI KHÔNG DÙNG Python 3.14 trở lên** vì hiện tại các thư viện C-extensions chưa hỗ trợ.
- 📥 **Tải về**: [Python 3.12.8 (Windows installer 64-bit)](https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe)
- ⚠️ **Khi cài đặt**: Hãy tích chọn vào ô **`Add python.exe to PATH`** ở màn hình đầu tiên trước khi nhấn *Install Now*.

### 1.2. Git for Windows
- 📥 **Tải về**: [Git for Windows](https://git-scm.com/download/win) (cứ bấm *Next* theo mặc định khi cài đặt).

### 1.3. Node.js (Khuyến nghị - phục vụ tải ảnh Zalo)
- Dùng để vận hành service tải ảnh tự động từ các nhóm Zalo (`zalo-service`).
- 📥 **Tải về**: [Node.js LTS (v20 hoặc v22)](https://nodejs.org/)
- *(Lưu ý: Nếu chưa cài Node.js, bạn vẫn dùng được 100% các tính năng Quét mặt, Xử lý ảnh, PDF, Excel; chỉ riêng tab tải ảnh Zalo sẽ tạm tắt).*

### 1.4. Trình duyệt Google Chrome hoặc Microsoft Edge
- Dùng để hiển thị giao diện web và hỗ trợ cơ chế quét ảnh Zalo qua trình duyệt tự động. (Microsoft Edge đã có sẵn trên Windows 10/11).

### 1.5. Microsoft Visual C++ Redistributable (Nếu máy mới tinh)
- Thường các thư viện như OpenCV và PyTorch yêu cầu bộ thư viện runtime này.
- 📥 **Tải về**: [VC_redist.x64.exe](https://aka.ms/vs/17/release/vc_redist.x64.exe)

---

## 2. 🚀 CÁC BƯỚC THỰC HIỆN SAU KHI CLONE DỰ ÁN

### Bước 1: Clone mã nguồn về máy tính
Mở **Terminal** (CMD hoặc PowerShell) và gõ:
```bash
git clone https://github.com/TysonNg/Chamconghinhanh.git
cd Chamconghinhanh
```

---

### Bước 2: Khởi chạy phần mềm

Bạn có thể lựa chọn 1 trong 2 cách sau:

### 👉 CÁCH 1: KHỞI CHẠY TỰ ĐỘNG 1-CLICK (KHUYÊN DÙNG)
Nhấp đúp chuột vào file:
```
PhanMemQuetMat.bat
```
Script thông minh sẽ tự động làm toàn bộ các việc sau:
1. Tìm và kiểm tra phiên bản Python phù hợp (Python 3.12).
2. Tự động tạo môi trường ảo cách ly `.venv` (tránh xung đột thư viện với các dự án khác trên máy).
3. Tự động cài đặt đầy đủ tất cả các thư viện từ `requirements.txt` vào `.venv`.
4. Tự động kiểm tra và cài đặt dependencies cho `zalo-service` (nếu máy có Node.js).
5. Khởi động server Flask nội bộ và **tự động mở trình duyệt web** tại:
   👉 **`http://127.0.0.1:5000`**

*(Lưu ý: Lần đầu tiên chạy script có thể mất từ 3 - 5 phút để tải và cài đặt các thư viện AI nặng như PyTorch, TensorFlow, OpenCV. Các lần chạy tiếp theo sẽ mở lên ngay lập tức sau 2 - 3 giây).*

---

### 👉 CÁCH 2: CÀI ĐẶT THỦ CÔNG QUA DÒNG LỆNH
Nếu bạn muốn tự quản lý môi trường qua Terminal / Command Prompt:

```bash
# 1. Di chuyển vào thư mục dự án
cd "d:\Projects\phan mem quet mat"

# 2. Khởi tạo môi trường ảo với Python 3.12
py -3.12 -m venv .venv

# 3. Kích hoạt môi trường ảo
# Trên Command Prompt (CMD):
.venv\Scripts\activate.bat
# Hoặc trên PowerShell:
# .venv\Scripts\Activate.ps1

# 4. Nâng cấp pip và cài đặt toàn bộ thư viện
python -m pip install --upgrade pip
pip install -r requirements.txt

# 5. Cài đặt thư viện cho Zalo Service (nếu có Node.js)
cd zalo-service
npm install
cd ..

# 6. Khởi động ứng dụng
python main.py
```
Sau đó mở trình duyệt web bất kỳ và truy cập địa chỉ: `http://127.0.0.1:5000`

---

## 3. 📂 CẤU TRÚC DỮ LIỆU CẦN NẮM ĐƯỢC

Dự án tự động khởi tạo cấu trúc thư mục làm việc khi khởi động. Các thư mục chính gồm:

| Thư mục | Chức năng | Ghi chú |
| :--- | :--- | :--- |
| `Ảnh BV/<Tên_Dự_Án>/` | Lưu ảnh chân dung nhân viên đối chiếu | Đặt ảnh chân dung rõ mặt để hệ thống học và so khớp |
| `input_images/<Tên_Dự_Án>/<YYYY-MM-DD>/` | Chứa ảnh camera chấm công hàng ngày | Sắp xếp theo ngày chụp định dạng Năm-Tháng-Ngày |
| `results/` | Chứa các file báo cáo Excel xuất ra | File chấm công đối chiếu, tỷ lệ nhận diện |
| `data/` | Database SQLite quản lý danh tính (`identity.sqlite3`) | Tự động tạo và lưu trữ mã nhân viên, dự án |
| `supplement_data/` | Lưu trữ hồ sơ bổ sung ảnh chấm công | Quản lý các batch duyệt ảnh bổ sung |
| `templates/word/` | Chứa mẫu file Word báo cáo giải trình (`.docx`) | Đã có sẵn file mẫu trong repo |
| `zalo-service/` | Server Node.js tải ảnh Zalo tự động | Chạy cổng ngầm 3001 |

---

## 4. 🧠 LƯU Ý KHI CHẠY LẦN ĐẦU TIÊN (TỰ ĐỘNG TẢI MODEL AI)

Trong lần đầu tiên bạn bấm chức năng quét mặt hoặc đọc watermark ngày giờ:
1. **Model DeepFace**: Sẽ tự động tải file weights (trọng số mạng nơ-ron nhận diện khuôn mặt) về thư mục:
   `C:\Users\<Tên_Máy_Tính>\.deepface\weights\` (~100MB - 500MB).
2. **Model EasyOCR**: Sẽ tự động tải model nhận diện ký tự tiếng Việt (`vi`) và tiếng Anh (`en`) về thư mục:
   `C:\Users\<Tên_Máy_Tính>\.EasyOCR\model\` (~100MB).

👉 **Vì vậy, hãy đảm bảo máy tính có kết nối Internet ổn định trong lần quét thử nghiệm đầu tiên.**

---

## 5. ❓ XỬ LÝ CÁC SỰ CỐ THƯỜNG GẶP (TROUBLESHOOTING)

### 🔴 Lỗi 1: "Python 3.14 không tương thích với NumPy/TensorFlow..."
- **Nguyên nhân**: Máy bạn đang cài Python 3.14 làm phiên bản mặc định.
- **Cách khắc phục**: Tải và cài đặt **Python 3.12 (64-bit)**. Script `PhanMemQuetMat.bat` sẽ tự động ưu tiên gọi đúng `py -3.12`.

### 🔴 Lỗi 2: PowerShell báo lỗi "Execution of scripts is disabled on this system"
- **Nguyên nhân**: Chính sách bảo mật của Windows PowerShell chặn chạy script `activate.ps1`.
- **Cách khắc phục**: Mở PowerShell với quyền Administrator và gõ:
  ```powershell
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```
  Hoặc chỉ cần dùng trực tiếp file `PhanMemQuetMat.bat` (chạy qua CMD không bị giới hạn này).

### 🔴 Lỗi 3: Port 5000 hoặc Port 3001 bị báo đang bận (Address already in use)
- **Nguyên nhân**: Lần chạy trước tắt chưa hết hoặc có ứng dụng khác đang chiếm port.
- **Cách khắc phục**:
  - `PhanMemQuetMat.bat` và `main.py` đã tích hợp cơ chế tự động giải phóng port 3001.
  - Đối với port 5000, mở CMD và gõ lệnh sau để đóng:
    ```cmd
    netstat -ano | findstr :5000
    taskkill /F /PID <PID_tìm_thấy>
    ```

### 🔴 Lỗi 4: Không nhận diện được watermark hoặc báo lỗi OCR
- Phần mềm mặc định sử dụng `EasyOCR` (tích hợp sẵn qua pip).
- Nếu muốn sử dụng thêm engine OCR truyền thống Tesseract:
  1. Tải bộ cài Tesseract cho Windows tại: [UB-Mannheim Tesseract](https://github.com/UB-Mannheim/tesseract/wiki)
  2. Cài vào thư mục: `C:\Program Files\Tesseract-OCR\` (nhớ tích chọn gói ngôn ngữ Vietnamese khi cài).

---

## 6. 🛑 CÁCH DỪNG PHẦN MỀM

- Nhấn tổ hợp phím **`Ctrl + C`** trong cửa sổ Command Prompt đang chạy, hoặc đơn giản là **đóng cửa sổ console** màu đen của phần mềm.
- Ứng dụng đã tích hợp Windows Job Object tự động giải phóng toàn bộ tiến trình con (Python Flask và Node.js Zalo Service) để không bị chạy ngầm làm nặng máy.
