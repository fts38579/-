@echo off
setlocal
title Kawauso Manager Kinako - Build EXE
cd /d "%~dp0"

echo.
echo ============================================================
echo   Kawauso Manager Kinako - Build Start
echo ============================================================
echo.

echo [Step 0] Installing required libraries...
py -m pip install pyinstaller tkcalendar --quiet
if %ERRORLEVEL% neq 0 ( echo [ERROR] pip install failed & pause & exit /b 1 )
echo [OK] Libraries ready.
echo.

echo [Step 1/1] Building EXE...
py -m PyInstaller --onefile --windowed ^
  --name "kawauso-kinako" ^
  --distpath ".." ^
  --workpath "build" ^
  --specpath "." ^
  --hidden-import selenium.webdriver.chrome.webdriver ^
  --hidden-import selenium.webdriver.chrome.service ^
  --hidden-import selenium.webdriver.chrome.options ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --collect-all TikTokLive ^
  --collect-all selenium ^
  --collect-all bs4 ^
  --collect-all matplotlib ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  app.py
if %ERRORLEVEL% neq 0 ( echo [ERROR] Build failed & pause & exit /b 1 )
echo [OK] Step 1/1 done.
echo.

echo ============================================================
echo   Build complete!
echo   Output: kawauso-kinako.exe (in project root folder)
echo ============================================================
pause
endlocal
