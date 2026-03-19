@echo off
chcp 65001 > nul
setlocal
title カワウソマネージャー きなこ - EXE ビルド
cd /d "%~dp0"

echo.
echo ============================================================
echo   カワウソマネージャー きなこ  EXE ビルド開始
echo ============================================================
echo.

echo [準備] PyInstaller をインストール中...
py -m pip install pyinstaller --quiet
if %ERRORLEVEL% neq 0 ( echo ❌ 失敗 & pause & exit /b 1 )
echo ✅ 準備完了
echo.

:: ── ① カワウソマネージャー きなこ.exe ────────────────────
echo [1/4] カワウソマネージャー きなこ.exe をビルド中...
py -m PyInstaller --onefile --console ^
  --name "カワウソマネージャー きなこ" ^
  --distpath ".." ^
  --workpath "build" ^
  --specpath "." ^
  --hidden-import selenium.webdriver.chrome.webdriver ^
  --hidden-import selenium.webdriver.chrome.service ^
  --hidden-import selenium.webdriver.chrome.options ^
  --collect-all TikTokLive ^
  --collect-all selenium ^
  --collect-all bs4 ^
  main.py
if %ERRORLEVEL% neq 0 ( echo ❌ 失敗 & pause & exit /b 1 )
echo ✅ [1/4] 完了
echo.

:: ── ② きなこのレポート.exe ───────────────────────────────
echo [2/4] きなこのレポート.exe をビルド中...
py -m pip install tkcalendar --quiet
py -m PyInstaller --onefile --windowed ^
  --name "きなこのレポート" ^
  --distpath ".." ^
  --workpath "build" ^
  --specpath "." ^
  --collect-all matplotlib ^
  --collect-all tkcalendar ^
  --collect-all babel ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  きなこのレポート.py
if %ERRORLEVEL% neq 0 ( echo ❌ 失敗 & pause & exit /b 1 )
echo ✅ [2/4] 完了
echo.

:: ── ③ 初期セットアップ.exe ──────────────────────────────
echo [3/4] 初期セットアップ.exe をビルド中...
py -m PyInstaller --onefile --windowed ^
  --name "初期セットアップ" ^
  --distpath ".." ^
  --workpath "build" ^
  --specpath "." ^
  "..\セットアップ\初期セットアップ.py"
if %ERRORLEVEL% neq 0 ( echo ❌ 失敗 & pause & exit /b 1 )
echo ✅ [3/4] 完了
echo.

:: ── ④ インサイト手動取得.exe ─────────────────────────────
echo [4/4] インサイト手動取得.exe をビルド中...
py -m PyInstaller --onefile --windowed ^
  --name "インサイト手動取得" ^
  --distpath ".." ^
  --workpath "build" ^
  --specpath "." ^
  --hidden-import selenium.webdriver.chrome.webdriver ^
  --hidden-import selenium.webdriver.chrome.service ^
  --hidden-import selenium.webdriver.chrome.options ^
  --collect-all selenium ^
  --collect-all bs4 ^
  --hidden-import pandas ^
  "インサイト手動取得.py"
if %ERRORLEVEL% neq 0 ( echo ❌ [4/4] 失敗 & pause & exit /b 1 )
echo ✅ [4/4] 完了
echo.

echo ============================================================
echo   ✅ 全EXE完了！プロジェクトフォルダ直下に生成されました
echo ============================================================
pause
endlocal
