@echo off
REM This batch file is a simple launcher for the JisuiArc2PDF.py script.
REM It passes all command-line arguments directly to the Python script.
REM The Python script itself will handle the interactive mode if no arguments are given,
REM or if you drag-and-drop a file onto this batch file.

cd /d "%~dp0"

REM Check for python executable
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: 'python' command not found.
    echo Please install Python and ensure it is in your system's PATH.
    pause
    goto :eof
)

REM Run the python script, passing all arguments
python "%~dp0\JisuiArc2PDF.py" %*


echo.
echo Done. Press any key to exit.
pause >nul