@echo off
setlocal enabledelayedexpansion

REM Set the text color to red (0C: black background, red text)
color 0C

echo ============================
echo 🔥 VocalScrub Clean Uninstall 🔥
echo ============================

REM Save current directory (folder where this .bat is located)
set "CURRENT_DIR=%~dp0"

echo.
echo 📦 Step 1: Uninstalling all pip packages (Python stays safe)...
pip freeze > "%CURRENT_DIR%piplist.txt"
for /F "delims=" %%i in (%CURRENT_DIR%piplist.txt) do (
    echo Uninstalling %%i ...
    pip uninstall -y %%i
)
del "%CURRENT_DIR%piplist.txt"

echo.
echo 🧹 Step 2: Cleaning pip cache...
pip cache purge

echo.

echo.
echo 💀 Step 4: Killing VocalScrub.exe if running...
taskkill /f /im VocalScrub.exe >nul 2>&1

echo.
echo 🗂️ Step 5: Preparing to delete VocalScrub folder...

REM Create a temporary self-deletion script
set "DELETE_SCRIPT=%temp%\delete_me.bat"
echo @echo off > "%DELETE_SCRIPT%"
echo ping 127.0.0.1 -n 3 >nul >> "%DELETE_SCRIPT%"
echo rmdir /s /q "%CURRENT_DIR%" >> "%DELETE_SCRIPT%"
echo del "%%~f0" >> "%DELETE_SCRIPT%"
echo exit >> "%DELETE_SCRIPT%"

echo.
echo 🧨 Step 6: Self-destructing script...
call "%DELETE_SCRIPT%"

exit
