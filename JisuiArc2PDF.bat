@echo off
REM �X�N���v�g�̂���ꏊ�Ɉړ�
cd /d "%~dp0"

REM PowerShell�X�N���v�g��Θb���[�h�ŋN��
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0\JisuiArc2PDF.ps1"

echo.
echo �������������܂����B�����L�[�������ƃE�B���h�E����܂��B
pause >nul