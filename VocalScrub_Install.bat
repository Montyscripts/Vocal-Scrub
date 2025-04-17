@echo off
SETLOCAL ENABLEEXTENSIONS
cd /d "%~dp0"

:: ===============================
:: Set Text Color to Lime Green
:: ===============================
color 0A

:: ===============================
:: Display Coffee Art
:: ===============================
echo       ( (
echo        ) )
echo     ........
echo     ^|      ^|]   ☕ VOCALSCRUB INSTALLER ☕
echo     \      /
echo      `----'     Brewing voice control...
echo                 Heating up vocal beans...
echo                 Loading fresh commands...
echo                 Please wait.
echo.
timeout /t 2 >nul

:: ===============================
:: Auto-run as administrator
:: ===============================
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "cmd.exe", "/c cd ""%~sdp0"" && %~s0", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /b
)

echo ########################################################
echo #           VocalScrub Automatic Installer             #
echo # This will automatically:                             #
echo # 1. Install all required packages                     #
echo # 2. Create the VocalScrub executable                  #
echo #                                                      #
echo # Please wait while the installer runs...              #
echo ########################################################

:: ===============================
:: Check internet connection
:: ===============================
echo Checking internet connection...
ping -n 2 google.com >nul
if errorlevel 1 (
    echo Error: No internet connection detected.
    pause
    exit /b 1
)

:: ===============================
:: Check for Python installation
:: ===============================
where python >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed. Please install Python from:
    echo https://www.python.org/downloads/
    echo.
    echo Follow the instructions on the Python website to install Python.
    pause
    exit /b 1
)

:: ===============================
:: Upgrade pip, setuptools, and wheel
:: ===============================
echo Upgrading pip, setuptools, and wheel...
python -m pip install --upgrade pip setuptools wheel

:: ===============================
:: Install PyInstaller early
:: ===============================
echo Installing PyInstaller...
python -m pip install pyinstaller

:: ===============================
:: Install PyAudio
:: ===============================
echo Installing PyAudio...
python -m pip install pyaudio

:: ===============================
:: Install required packages
:: ===============================
echo Installing required packages...
python -m pip install -r requirements.txt
if %errorlevel% NEQ 0 (
    echo Failed to install Python packages. Trying manual Pillow installation...

    :: Trying to install Pillow from source
    python -m pip install pillow==9.5.0 --no-binary :all:

    if %errorlevel% NEQ 0 (
        echo ERROR: Failed to install Pillow manually using the source.
        pause
        exit /b 1
    )
)

:: ===============================
:: Create VocalScrub executable
:: ===============================
echo Creating VocalScrub executable...
echo This may take a few minutes...
pyinstaller --onefile --noconsole --icon=Icon.png --add-data "Button.mp3;." --add-data "Click.mp3;." --add-data "Hover.mp3;." --add-data "Icon.png;." --add-data "Wallpaper.png;." --add-data "Button.png;." VocalScrub.py

:: ===============================
:: Check if executable was created
:: ===============================
if not exist "dist\VocalScrub.exe" (
    echo ERROR: Executable creation failed.
    pause
    exit /b 1
)

:: ===============================
:: Done!
:: ===============================
echo.
echo ########################################################
echo #         Installation completed successfully! 🚀      #
echo ########################################################
echo Installation is complete. You can close this window now.
pause
exit /b
