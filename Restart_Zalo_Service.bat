@echo off
chcp 65001 > nul
title Khởi Động Lại Zalo Service (Port 3001)

echo.
echo ==============================================================
echo        KHỞI ĐỘNG LẠI RIÊNG DỊCH VỤ ZALO (PORT 3001)
echo ==============================================================
echo.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%zalo-service"

echo [1/3] Đang dọn dẹp tiến trình cũ trên port 3001...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :3001 ^| findstr LISTENING') do (
    echo   - Đóng PID %%a
    taskkill /F /PID %%a > nul 2>&1
)

echo [2/3] Kiểm tra môi trường Node.js...
where node > nul 2>&1
if errorlevel 1 (
    echo.
    echo [LỖI] Không tìm thấy Node.js! Vui lòng cài Node.js tại https://nodejs.org
    echo.
    pause
    exit /b 1
)

echo [3/3] Đang khởi động Zalo Service (node server.js)...
echo.
echo Dịch vụ Zalo sẽ chạy trên port 3001.
echo Giữ cửa sổ này để xem log trực tiếp, hoặc nhấn Ctrl+C để dừng.
echo --------------------------------------------------------------
echo.

node server.js
pause
