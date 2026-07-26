@echo off
chcp 65001 >nul
:menu
cls
echo =======================================================
echo     TỰ ĐỘNG HÓA WORKFORCE DASHBOARD (CHỌN BƯỚC)
echo =======================================================
echo 1. Tải dữ liệu từ MISA (Bước 1)
echo 2. Làm mới PowerQuery (Bước 2)
echo 3. Chạy ETL tạo data.json (Bước 3)
echo 4. Deploy lên GitHub Pages (Bước 4)
echo 5. Chạy TOÀN BỘ các bước (1-4)
echo 6. Thoát
echo =======================================================
set choice=
set /p choice="Chon so (1-6): "

set step=all
if "%choice%"=="1" set step=1
if "%choice%"=="2" set step=2
if "%choice%"=="3" set step=3
if "%choice%"=="4" set step=4
if "%choice%"=="5" set step=all
if "%choice%"=="6" exit

echo.
echo Khởi chạy...
python "G:\My Drive\Dữ liệu nhân sự\.agents\skills\PC_attendance_management\scripts\sync_and_deploy.py" --step %step%
if errorlevel 1 (
    "C:\Python312\python.exe" "G:\My Drive\Dữ liệu nhân sự\.agents\skills\PC_attendance_management\scripts\sync_and_deploy.py" --step %step%
)

echo.
pause
goto menu
