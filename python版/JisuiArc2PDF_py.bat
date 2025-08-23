
@echo off
setlocal

REM ============================================================================
REM Argument Parsing for Python Script
REM ============================================================================
set "log_path_param="
set "file_args_temp="

:ParseArgs
if "%~1"=="" goto EndParseArgs
if /i "%~1"=="--LogPath" (
    set "log_path_param=--LogPath ""%~2"""
    shift
    shift
    goto ParseArgs
)
set "file_args_temp=%file_args_temp% %1"
shift
goto ParseArgs

:EndParseArgs
if defined file_args_temp ( set "file_args=%file_args_temp:~1%" ) else ( set "file_args=" )

if not defined file_args (
    echo ERROR: No input files specified.
    echo Usage: %~n0 "*.rar"
    echo        %~n0 "MyBook.zip" --LogPath "C:\My Logs"
    pause
    goto :eof
)

REM ============================================================================
REM Interactive Parameter Prompt for Python Script
REM ============================================================================
set "py_params="

set /p skip_in="Skip compression (y/n)? [n]: "
if /i "%skip_in%"=="y" (
    set "py_params=%py_params% -sc"
    goto :execute_command
)

set /p quality_in="Quality (1-100) [85]: "
if not "%quality_in%"=="" set "py_params=%py_params% -q %quality_in%"

set /p sat_in="Saturation threshold [0.05]: "
if not "%sat_in%"=="" set "py_params=%py_params% -s %sat_in%"

set /p tcr_in="Total compression threshold (0-100, optional): "
if not "%tcr_in%"=="" set "py_params=%py_params% -tcr %tcr_in%"

set /p deskew_in="Deskew (auto-straighten) (y/n)? [n]: "
if /i "%deskew_in%"=="y" (
    set "py_params=%py_params% -ds"
)

REM Ask for contrast
set /p gray_contrast_in="Adjust Grayscale contrast (y/n)? [n]: "
if /i "%gray_contrast_in%"=="y" (
    set /p level_in="Grayscale Level value [10%%,90%%]: "
    if "%level_in%"=="" set "level_in=10%%,90%%"
    set "py_params=%py_params% -gl ""%level_in%"""
)

set "auto_contrast_in="
set /p auto_contrast_in="Auto-adjust Color contrast (y/n)? [n]: "
if /i "%auto_contrast_in%"=="y" (
    set "py_params=%py_params% -ac"
) else (
    set /p color_contrast_in="Adjust Color contrast (y/n)? [n]: "
    if /i "%color_contrast_in%"=="y" (
        set /p bright_in="Color Brightness-Contrast value [0x25]: "
        if "%bright_in%"=="" set "bright_in=0x25"
        set "py_params=%py_params% -cc ""%bright_in%"""
    )
)

set /p trim_in="Trim margins (y/n)? [n]: "
if /i "%trim_in%"=="y" (
    set "py_params=%py_params% -t"
    set /p fuzz_in="Fuzz factor for trim (e.g., 1%%) [1%%]: "
    if not "%fuzz_in%"=="" (
        set "py_params=%py_params% --Fuzz ""%fuzz_in%"""
    )
)

set /p linearize_in="Linearize PDF (web optimization) (y/n)? [n]: "
if /i "%linearize_in%"=="y" (
    set "py_params=%py_params% -lin"
)

REM ============================================================================
REM Resolution Settings
REM ============================================================================
set /p res_choice="Resolution: 1=Height, 2=Paper+DPI [2]: "
if "%res_choice%"=="1" goto :ask_height
goto :ask_paper_dpi

:ask_height
set /p height_in="Height (pixels): "
if not "%height_in%"=="" set "py_params=%py_params% --Height %height_in%"
set /p dpi_for_h_in="DPI [144]: "
if not "%dpi_for_h_in%"=="" set "py_params=%py_params% --Dpi %dpi_for_h_in%"
goto :execute_command

:ask_paper_dpi
set /p paper_in="Paper Size [A4]: "
if "%paper_in%"=="" set "paper_in=A4"
set /p dpi_in="DPI [144]: "
if "%dpi_in%"=="" set "dpi_in=144"
set "py_params=%py_params% --PaperSize %paper_in% --Dpi %dpi_in%"

REM ============================================================================
REM Command Execution
REM ============================================================================
:execute_command
REM Check for python executable
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: 'python' command not found.
    echo Please install Python and ensure it is in your system's PATH.
    pause
    goto :eof
)

set "final_command=python "%~dp0\JisuiArc2PDF.py" %file_args% %py_params% %log_path_param%"

echo.
echo Running: %final_command%
echo.
set /p "confirm=Press Enter to run, or Ctrl+C to cancel..."

%final_command%

REM ============================================================================
REM Exit
REM ============================================================================
echo.
echo Done. Press any key to exit.
pause >nul
