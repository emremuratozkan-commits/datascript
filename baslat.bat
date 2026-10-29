@echo off
title LinkedIn Otomasyon Baslatici
cd /d "%~dp0"

echo ========================================================
echo Eski calisan surecler temizleniyor...
echo ========================================================
for /f "tokens=5" %%a in ('netstat -aon ^| find ":8501" ^| find "LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)

echo ========================================================
echo LinkedIn Otomasyon Arayuzu Baslatiliyor...
echo Web adresi: http://localhost:8501
echo ========================================================
echo.

start "" "http://localhost:8501"
python -m streamlit run app.py --server.port 8501

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Bir hata olustu. Hata detayini gormek icin lutfen bir tusa basin...
    pause
)
