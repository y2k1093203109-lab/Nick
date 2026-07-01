@echo off
chcp 65001 >nul
title EU MDR PSUR 審查系統啟動器

echo ===================================================
echo   正在啟動 PSUR 自動化與視覺審查系統，請稍候...
echo ===================================================
echo.

pushd "%~dp0"

C:\Users\P0905\AppData\Local\Programs\Python\Python312\python.exe -m streamlit run automate_psur.py

popd
pause
