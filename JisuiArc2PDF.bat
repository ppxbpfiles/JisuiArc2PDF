@echo off
chcp 65001 >nul
REM �X�N���v�g�̂���ꏊ�Ɉړ�
cd /d "%~dp0"

REM PowerShell�X�N���v�g�����s (���������ׂēn��)
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0\JisuiArc2PDF.ps1" %*

echo.
echo �������������܂����B�����L�[�������ƃE�B���h�E����܂��B
pause >nul
