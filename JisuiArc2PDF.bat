@echo off
REM スクリプトのある場所に移動
cd /d "%~dp0"

REM PowerShellスクリプトを対話モードで起動
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0\JisuiArc2PDF.ps1"

echo.
echo 処理が完了しました。何かキーを押すとウィンドウを閉じます。
pause >nul