@echo off
cd /d "%~dp0"
title TGS 3 - Klasifikasi Email Spam
echo ======================================================
echo   TGS 3 - Klasifikasi Email Spam (Naive Bayes)
echo   Data Mining dan Information Retrieval - IFB-307
echo ======================================================
echo.

where python >nul 2>nul
if %errorlevel%==0 (
    python klasifikasi_spam.py --menu
) else (
    py klasifikasi_spam.py --menu
)

if errorlevel 1 (
    echo.
    echo [ERROR] Program gagal dijalankan. Pastikan Python dan
    echo         scikit-learn sudah terpasang:
    echo         python -m pip install scikit-learn pandas matplotlib
)
echo.
echo Tekan ENTER untuk menutup jendela...
set /p dummy=""
