@echo off
chcp 65001 >nul
REM スクリプトのある場所に移動
cd /d "%~dp0"

REM PowerShellスクリプトを実行 (引数をすべて渡す)
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0\JisuiArc2PDF.ps1" %*

echo.
echo 処理が完了しました。何かキーを押すとウィンドウを閉じます。
pause >nul
