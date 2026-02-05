@echo off
chcp 65001 > nul
title Phần Mềm Nhận Diện Khuôn Mặt & Chấm Công

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║     PHẦN MỀM NHẬN DIỆN KHUÔN MẶT ^& CHẤM CÔNG                ║
echo ║     Đang khởi động...                                        ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

REM Kiểm tra Python
python --version > nul 2>&1
if errorlevel 1 (
    py --version > nul 2>&1
    if errorlevel 1 (
        echo [LỖI] Không tìm thấy Python! Vui lòng cài đặt Python trước.
        echo Tải tại: https://www.python.org/downloads/
        pause
        exit /b 1
    ) else (
        set PYTHON_CMD=py
    )
) else (
    set PYTHON_CMD=python
)

REM Cài đặt dependencies nếu chưa có
echo Đang kiểm tra dependencies...
%PYTHON_CMD% -c "import flask" > nul 2>&1
if errorlevel 1 (
    echo Đang cài đặt dependencies...
    %PYTHON_CMD% -m pip install -r requirements.txt -q
)

echo Đang khởi động server...
echo.
echo ═══════════════════════════════════════════════════════════════
echo   Trình duyệt sẽ tự động mở. Nếu không, hãy truy cập:
echo   http://127.0.0.1:5000
echo.
echo   Để dừng phần mềm, đóng cửa sổ này hoặc nhấn Ctrl+C
echo ═══════════════════════════════════════════════════════════════
echo.

%PYTHON_CMD% main.py
pause
