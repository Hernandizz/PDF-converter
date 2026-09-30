@echo off
title HD Word to PDF Converter Pro
color 0b

echo ================================================================
echo       HD WORD TO PDF CONVERTER PRO (LOSSLESS PRINT QUALITY)
echo ================================================================
echo.
echo Memeriksa file dependensi...
python -m pip install -r requirements.txt --quiet

echo.
echo Menjalankan web server lokal...
start "" http://127.0.0.1:5000
python app.py
pause
