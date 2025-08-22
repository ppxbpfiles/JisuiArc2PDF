@echo off
setlocal

REM --- Argument Parsing ---
set "log_path_param="
set "file_args_temp="
:ParseArgs
if "%~1"=="" goto EndParseArgs
if /i "%~1"=="-LogPath" (
    set "log_path_param=-LogPath ""%~2"""
    shift
    shift
    goto ParseArgs
)
set "file_args_temp=%file_args_temp% %1"
shift
goto ParseArgs
:EndParseArgs

if defined file_args_temp ( set "file_args=%file_args_temp:~1%" ) else ( set "file_args=" )

REM --- Check for Input Files ---
if not defined file_args (
    echo ERROR: No input files specified.
    echo Usage: %~n0 "*.rar"
    echo        %~n0 "MyBook.zip" -LogPath "C:\My Logs"
    pause
    goto :eof
)

REM --- Build PowerShell Command ---
set "ps_params="

set /p skip_in="Skip compression (y/n)? [n]: "
if /i "%skip_in%"=="y" (
    set "ps_params=%ps_params% -SkipCompression"
    goto :execute_command
)

set /p quality_in="Quality (1-100) [85]: "
if not "%quality_in%"=="" set "ps_params=%ps_params% -Quality %quality_in%"

set /p sat_in="Saturation threshold [0.05]: "
if not "%sat_in%"=="" set "ps_params=%ps_params% -SaturationThreshold %sat_in%"

set /p mcr_in="Min compression ratio (optional): "
if not "%mcr_in%"=="" set "ps_params=%ps_params% -MinCompressionRatio %mcr_in%"

set /p res_choice="Resolution: 1=Height, 2=Paper+DPI [2]: "
if "%res_choice%"=="1" goto :ask_height
goto :ask_paper_dpi

:ask_height
set /p height_in="Height (pixels): "
if not "%height_in%"=="" set "ps_params=%ps_params% -Height %height_in%"
set /p dpi_for_h_in="DPI [144]: "
if not "%dpi_for_h_in%"=="" set "ps_params=%ps_params% -Dpi %dpi_for_h_in%"
goto :execute_command

:ask_paper_dpi
set /p paper_in="Paper Size [A4]: "
if "%paper_in%"=="" set "paper_in=A4"
set /p dpi_in="DPI [144]: "
if "%dpi_in%"=="" set "dpi_in=144"
set "ps_params=%ps_params% -PaperSize %paper_in% -Dpi %dpi_in%"

:execute_command
set "final_command=pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0\JisuiArc2PDF.ps1" %file_args% %ps_params% %log_path_param%"

echo.
echo Running: %final_command%
echo.
set /p "confirm=Press Enter to run, or Ctrl+C to cancel..."

%final_command%

echo.
echo Done. Press any key to exit.
pause >nul