@echo off
chcp 65001 > nul
setlocal
title カワウソマネージャー きなこ - EXE ビルド（統合版）
cd /d "%~dp0"

echo.
echo ============================================================
echo   カワウソマネージャー きなこ  統合EXE ビルド開始
echo ============================================================
echo.

echo [準備] 必要ライブラリをインストール中...
py -m pip install pyinstaller tkcalendar --quiet
if %ERRORLEVEL% neq 0 ( echo ❌ 失敗 & pause & exit /b 1 )
echo ✅ 準備完了
echo.

:: ── 統合アプリ（カワウソマネージャー きなこ.exe）────────────────
echo [1/1] カワウソマネージャー きなこ.exe をビルド中...
py -m PyInstaller --onefile --windowed ^
  --name "カワウソマネージャー きなこ" ^
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
if %ERRORLEVEL% neq 0 ( echo ❌ 失敗 & pause & exit /b 1 )
echo ✅ [1/1] 完了
echo.

echo ============================================================
echo   ✅ ビルド完了！
echo   プロジェクトフォルダ直下に
echo   「カワウソマネージャー きなこ.exe」が生成されました
echo ============================================================
pause
endlocal
