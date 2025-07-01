@echo off
title PracticeMojo Connection Tester

echo ====================================
echo PracticeMojo Connection Tester
echo ====================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python and add it to your PATH
    pause
    exit /b 1
)

:: Install required packages if missing
echo Checking required packages...
pip install requests beautifulsoup4 lxml >nul 2>&1

:: Run the tester
echo.
echo Running connection tests...
echo.
python PracticeMojoTester.py

:: Keep window open
echo.
pause