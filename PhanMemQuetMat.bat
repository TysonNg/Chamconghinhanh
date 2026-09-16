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
py -3.12 --version > nul 2>&1
if not errorlevel 1 (
    set PYTHON_CMD=py -3.12
) else (
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
)

REM Cài đặt dependencies nếu chưa có
echo Đang kiểm tra dependencies...
%PYTHON_CMD% -c "import flask, piexif" > nul 2>&1
if errorlevel 1 (
    echo Đang cài đặt dependencies...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [LỖI] Không thể cài đặt dependencies. Vui lòng kiểm tra kết nối mạng và chạy lại.
        pause
        exit /b 1
    )
    %PYTHON_CMD% -c "import flask, piexif" > nul 2>&1
    if errorlevel 1 (
        echo [LỖI] Dependencies vẫn chưa đầy đủ sau khi cài đặt.
        pause
        exit /b 1
    )
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
