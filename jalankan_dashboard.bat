@echo off
title PintarSaham Dashboard Local Server
echo ========================================================
echo   Memulai PintarSaham Dashboard di http://localhost:8080
echo ========================================================
echo.
timeout /t 1 >nul
start http://localhost:8080/dashboard.html
python -m http.server 8080
pause
