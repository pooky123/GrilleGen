@echo off
title Bingo Grid Generator — Setup
color 0A

echo.
echo  ============================================
echo   Bingo Grid Generator — First Time Setup
echo  ============================================
echo.
echo  This will install the required Python library
echo  (Pillow) needed to run the app.
echo.
echo  Make sure you have Python installed first.
echo  If you don't, go to: https://python.org/downloads
echo  and check "Add Python to PATH" during install.
echo.
pause

echo.
echo  Installing Pillow...
echo.

pip install Pillow

echo.
if %ERRORLEVEL% == 0 (
    echo  ============================================
    echo   SUCCESS! Pillow installed correctly.
    echo   You can now run bingo_generator.py
    echo  ============================================
) else (
    echo  ============================================
    echo   ERROR! Something went wrong.
    echo.
    echo   Most likely cause: Python is not added
    echo   to PATH. Re-run the Python installer,
    echo   click Modify, and enable "Add to PATH".
    echo  ============================================
)

echo.
pause
