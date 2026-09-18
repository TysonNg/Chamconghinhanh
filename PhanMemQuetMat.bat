@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion
title Phần Mềm Nhận Diện Khuôn Mặt & Chấm Công

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║     PHẦN MỀM NHẬN DIỆN KHUÔN MẶT ^& CHẤM CÔNG                ║
echo ║     Hệ thống khởi động tự động (Auto-Setup & Run)            ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM ==============================================================
REM BƯỚC 1: XÁC ĐỊNH TRÌNH THÔNG DỊCH PYTHON THÍCH HỢP
REM ==============================================================
echo [1/4] Đang kiểm tra môi trường Python...

set "SYS_PYTHON="

REM Ưu tiên 1: py -3.12
py -3.12 --version > nul 2>&1
if not errorlevel 1 (
    set "SYS_PYTHON=py -3.12"
    goto :python_found
)

REM Ưu tiên 2: py -3.11
py -3.11 --version > nul 2>&1
if not errorlevel 1 (
    set "SYS_PYTHON=py -3.11"
    goto :python_found
)

REM Ưu tiên 3: py -3.10
py -3.10 --version > nul 2>&1
if not errorlevel 1 (
    set "SYS_PYTHON=py -3.10"
    goto :python_found
)

REM Ưu tiên 4: lệnh python mặc định trong PATH
python --version > nul 2>&1
if not errorlevel 1 (
    set "SYS_PYTHON=python"
    goto :python_found
)

REM Ưu tiên 5: lệnh py launcher chung
py --version > nul 2>&1
if not errorlevel 1 (
    set "SYS_PYTHON=py"
    goto :python_found
)

REM Không tìm thấy Python
echo.
echo [LỖI] Không tìm thấy Python trên máy tính của bạn!
echo.
echo HƯỚNG DẪN CÀI ĐẶT:
echo 1. Tải Python 3.12 (64-bit) tại:
echo    https://www.python.org/downloads/release/python-3128/
echo 2. Khi chạy cài đặt, BẮT BUỘC TÍCH VÀO:
echo    "[X] Add python.exe to PATH"
echo 3. Sau khi cài xong, mở lại file PhanMemQuetMat.bat này.
echo.
pause
exit /b 1

:python_found
echo   - Phát hiện Python hệ thống: %SYS_PYTHON%

REM ==============================================================
REM BƯỚC 2: QUẢN LÝ MÔI TRƯỜNG ẢO (.venv)
REM ==============================================================
set "VENV_DIR=%SCRIPT_DIR%.venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    echo [2/4] Chưa có môi trường ảo .venv, đang tự động khởi tạo...
    echo   (Vui lòng đợi vài giây...)
    %SYS_PYTHON% -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo   [CẢNH BÁO] Không thể tạo .venv tự động. Sẽ sử dụng Python hệ thống.
        set "RUN_PYTHON=%SYS_PYTHON%"
    ) else (
        echo   - Khởi tạo .venv thành công tại: %VENV_DIR%
        set "RUN_PYTHON=%VENV_PYTHON%"
    )
) else (
    echo [2/4] Đã tìm thấy môi trường ảo: .venv
    set "RUN_PYTHON=%VENV_PYTHON%"
)

REM ==============================================================
REM BƯỚC 3: KIỂM TRA VÀ CÀI ĐẶT THƯ VIỆN PYTHON (requirements.txt)
REM ==============================================================
echo [3/4] Đang kiểm tra các thư viện cần thiết...

"%RUN_PYTHON%" -c "import flask, cv2, deepface, easyocr, fitz, xlrd, docx, requests" > nul 2>&1
if errorlevel 1 (
    echo   - Thiếu một số thư viện quan trọng. Đang tiến hành cài đặt từ requirements.txt...
    echo   - Quá trình này có thể mất 3-5 phút trong lần chạy đầu tiên. Vui lòng giữ kết nối mạng.
    echo.
    "%RUN_PYTHON%" -m pip install --upgrade pip
    "%RUN_PYTHON%" -m pip install -r "%SCRIPT_DIR%requirements.txt"
    if errorlevel 1 (
        echo.
        echo [LỖI] Cài đặt thư viện thất bại! Vui lòng kiểm tra kết nối mạng và thử lại.
        pause
        exit /b 1
    )
    echo   - Cài đặt thư viện Python hoàn tất!
) else (
    echo   - Các thư viện Python chính đã sẵn sàng.
)

REM ==============================================================
REM BƯỚC 4: KIỂM TRA ZALO SERVICE (NODE.JS)
REM ==============================================================
set "ZALO_DIR=%SCRIPT_DIR%zalo-service"
if exist "%ZALO_DIR%" (
    where node > nul 2>&1
    if errorlevel 1 (
        echo   [LƯU Ý] Chưa cài Node.js. Các tính năng quét mặt, PDF, Excel vẫn chạy bình thường.
        echo   (Nếu cần tải ảnh tự động từ Zalo, hãy cài Node.js tại: https://nodejs.org)
    ) else (
        if not exist "%ZALO_DIR%\node_modules" (
            echo   - Đang cài đặt thư viện Node.js cho Zalo Service...
            pushd "%ZALO_DIR%"
            call npm install --no-audit --no-fund
            popd
            echo   - Cài đặt Zalo Service hoàn tất.
        ) else (
            echo   - Zalo Service đã sẵn sàng.
        )
    )
)

REM ==============================================================
REM BƯỚC 5: KHỞI ĐỘNG PHẦN MỀM
REM ==============================================================
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║            ĐANG KHỞI ĐỘNG SERVER PHẦN MỀM...                 ║
echo ╠══════════════════════════════════════════════════════════════╣
echo ║  Trình duyệt sẽ tự động mở sau vài giây.                    ║
echo ║  Địa chỉ truy cập:  http://127.0.0.1:5000                   ║
echo ║                                                             ║
echo ║  Để dừng phần mềm: Đóng cửa sổ này hoặc nhấn Ctrl+C         ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

"%RUN_PYTHON%" "%SCRIPT_DIR%main.py"
pause
