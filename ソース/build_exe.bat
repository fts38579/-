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
py -m pip install pyinstaller PyQt6 pyqtgraph matplotlib pandas openpyxl ^
    selenium webdriver-manager TikTokLive beautifulsoup4 --quiet
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
  --hidden-import webdriver_manager ^
  --hidden-import webdriver_manager.chrome ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  --hidden-import PyQt6 ^
  --hidden-import PyQt6.QtWidgets ^
  --hidden-import PyQt6.QtCore ^
  --hidden-import PyQt6.QtGui ^
  --hidden-import matplotlib ^
  --collect-all TikTokLive ^
  --collect-all selenium ^
  --collect-all webdriver_manager ^
  --collect-all bs4 ^
  --collect-all matplotlib ^
  --collect-all PyQt6 ^
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
