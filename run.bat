@echo off
title HD Word to PDF Converter Pro
color 0b

echo ================================================================
echo       HD WORD TO PDF CONVERTER PRO (LOSSLESS PRINT QUALITY)
echo ================================================================
echo.
echo [1/2] Memeriksa dependensi sistem...
python -c "import flask, comtypes, docx, PIL, pypdf" >nul 2>nul
if %errorlevel% neq 0 (
    echo Menginstal dependensi yang belum ada...
    python -m pip install -r requirements.txt
) else (
    echo Dependensi lengkap. Fast boot aktif.
)

echo.
echo [2/2] Menjalankan web server lokal...
python app.py
if %errorlevel% neq 0 (
    echo.
    echo Terjadi kendala saat menjalankan server.
    pause
)
