@echo off
REM Prevents the commands from being displayed in the console.

setlocal
REM Enables localized environment variables. Changes are discarded when the script exits.

REM ============================================================================
REM Argument Parsing
REM ============================================================================
REM This section iterates through all command-line arguments provided to the script.
REM It specifically looks for the -LogPath parameter and separates it from the
REM file path arguments (like "*.rar" or "book.zip").

set "log_path_param="
set "file_args_temp="

:ParseArgs
REM If there are no more arguments, jump to the end of parsing.
if "%~1"=="" goto EndParseArgs

REM Check if the current argument is -LogPath (case-insensitive).
if /i "%~1"=="-LogPath" (
    REM If it is, store the parameter and its value, then skip to the next two arguments.
    set "log_path_param=-LogPath ""%~2"""
    shift
    shift
    goto ParseArgs
)

REM If it's not -LogPath, assume it's a file path and append it to a temporary variable.
set "file_args_temp=%file_args_temp% %1"
shift
goto ParseArgs

:EndParseArgs

REM Clean up the collected file arguments by removing the leading space.
if defined file_args_temp ( set "file_args=%file_args_temp:~1%" ) else ( set "file_args=" )


REM ============================================================================
REM Input File Check
REM ============================================================================
REM Checks if any file paths were provided. If not, displays an error and exits.
if not defined file_args (
    echo ERROR: No input files specified.
    echo Usage: %~n0 "*.rar"
    echo        %~n0 "MyBook.zip" -LogPath "C:\My Logs"
    pause
    goto :eof
)


REM ============================================================================
REM Interactive Parameter Prompt
REM ============================================================================
REM This section interactively asks the user for conversion settings.
REM These settings are used to build a string of parameters for the PowerShell script.

REM Initialize the parameter string.
set "ps_params="

REM Ask the user if they want to skip the compression/optimization steps.
set /p skip_in="Skip compression (y/n)? [n]: "
if /i "%skip_in%"=="y" (
    set "ps_params=%ps_params% -SkipCompression"
    goto :execute_command
)

REM Ask for JPEG quality.
set /p quality_in="Quality (1-100) [85]: "
if not "%quality_in%"=="" set "ps_params=%ps_params% -Quality %quality_in%"

REM Ask for the saturation threshold for grayscale detection.
set /p sat_in="Saturation threshold [0.05]: "
if not "%sat_in%"=="" set "ps_params=%ps_params% -SaturationThreshold %sat_in%"

REM Ask for the total compression ratio threshold.
set /p tcr_in="Total compression threshold (0-100, optional): "
if not "%tcr_in%"=="" set "ps_params=%ps_params% -TotalCompressionThreshold %tcr_in%"


REM ============================================================================
REM Resolution Settings
REM ============================================================================
REM Asks the user to choose between two methods for setting the output resolution.

set /p res_choice="Resolution: 1=Height, 2=Paper+DPI [2]: "
if "%res_choice%"=="1" goto :ask_height
goto :ask_paper_dpi

:ask_height
REM Method 1: Specify height in pixels directly.
set /p height_in="Height (pixels): "
if not "%height_in%"=="" set "ps_params=%ps_params% -Height %height_in%"
set /p dpi_for_h_in="DPI [144]: "
if not "%dpi_for_h_in%"=="" set "ps_params=%ps_params% -Dpi %dpi_for_h_in%"
goto :execute_command

:ask_paper_dpi
REM Method 2: Specify paper size and DPI to calculate height.
set /p paper_in="Paper Size [A4]: "
if "%paper_in%"=="" set "paper_in=A4"
set /p dpi_in="DPI [144]: "
if "%dpi_in%"=="" set "dpi_in=144"
set "ps_params=%ps_params% -PaperSize %paper_in% -Dpi %dpi_in%"


REM ============================================================================
REM Command Execution
REM ============================================================================
:execute_command
REM Assemble the final PowerShell command.
REM %~dp0 is the directory of the batch script itself.
REM %file_args% contains the input file paths.
REM %ps_params% contains the interactive settings.
REM %log_path_param% contains the -LogPath argument, if it was provided.
set "final_command=pwsh -NoProfile -ExecutionPolicy Bypass -File ""%~dp0\JisuiArc2PDF.ps1"" %file_args% %ps_params% %log_path_param%"

REM Display the command that will be run.
echo.
echo Running: %final_command%
echo.
REM Wait for user confirmation before executing.
set /p "confirm=Press Enter to run, or Ctrl+C to cancel..."

REM Execute the command.
%final_command%

REM ============================================================================
REM Exit
REM ============================================================================
echo.
echo Done. Press any key to exit.
pause >nul
