@echo off
setlocal enabledelayedexpansion
REM Get the directory of this script
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
REM ===================================================================
REM ==================== CONFIGURATION SECTION ========================
REM ===================================================================
REM Define all your scripts, icons, and paths here.
REM Build Toggles
REM Set to "YES" to compile, any other value to skip
set "COMPILE_WORMHOLE=YES"
REM Main Directories
set "OUTPUT_DIR=%SCRIPT_DIR%"
set "LOG_DIR=%SCRIPT_DIR%\Log"
REM Requirements File
set "REQUIREMENTS_FILE=%SCRIPT_DIR%\requirements.txt"
REM Script 1: Wormhole
set "WORMHOLE_SCRIPT=wormhole.py"
set "WORMHOLE_ICON=%SCRIPT_DIR%\Icons\Wormhole_Icon.ico"
set "WORMHOLE_BUILD_NAME=wormhole"
REM (Example) --add-data "path\to\file;destination\folder"
REM set "DATA_1=%SCRIPT_DIR%\assets\config.json;."
REM set "DATA_2=%SCRIPT_DIR%\assets\images;images"
REM ===================================================================
REM ================== SCRIPT EXECUTION (No Need to Edit) =============
REM ===================================================================
REM 1. Dependency Checks
echo Checking dependencies...
REM Use 'py -m pip' to ensure we're using the launcher's pip
py -m pip install pyinstaller
if %errorlevel% neq 0 (
    echo ▩▩▩ ERROR: Failed to install PyInstaller. ▩▩▩
    pause
    goto :eof
)
REM Ensure pywin32 is installed for single-instance logic
py -m pip show pywin32 >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing pywin32 for single-instance enforcement...
    py -m pip install pywin32
    if %errorlevel% neq 0 (
        echo ▩▩▩ ERROR: Failed to install pywin32. ▩▩▩
        pause
        goto :eof
    )
)
REM 2. Install dependencies from requirements.txt
echo.
echo Installing script dependencies from %REQUIREMENTS_FILE%...
if not exist "%REQUIREMENTS_FILE%" (
    echo ▩▩▩ ERROR: %REQUIREMENTS_FILE% not found! ▩▩▩
    echo Please create it in the script directory and add your project's dependencies:
    echo e.g., customtkinter, cryptography, python-docx, odfpy, Pillow, pywin32
    pause
    goto :eof
)
REM Use 'py -m pip' to install from requirements
py -m pip install -r "%REQUIREMENTS_FILE%"
if %errorlevel% neq 0 (
    echo ▩▩▩ ERROR: Failed to install dependencies from requirements.txt. ▩▩▩
    echo ▩▩▩ Check the output above for errors. ▩▩▩
    pause
    goto :eof
)
echo Dependencies installed.
echo.
REM 3. Setup Directories
echo Setting up directories...
if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"
REM 4. Clean Old Build Artifacts
echo Cleaning old build artifacts from %LOG_DIR%...
if exist "%LOG_DIR%\build" rmdir /s /q "%LOG_DIR%\build"
del /q "%LOG_DIR%\*.spec" 2>nul
del /q "%LOG_DIR%\*.log" 2>nul
del /q "%LOG_DIR%\*.sln" 2>nul
REM 5. Process requirements.txt for hidden imports
echo ---
echo Setting up hidden imports for PyInstaller...
set "HIDDEN_IMPORTS="
REM Add all required hidden imports based on your list.
REM Standard libraries like os, sys, json, etc., are usually found automatically.
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=customtkinter"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=tkinter"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=PIL"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=PIL.ImageTk"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=reportlab"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=pypdf"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=py7zr"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=docx"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=bs4"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=pptx"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=openpyxl"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=ezodf"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=odf"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=odf.opendocument"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=odf.text"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=striprtf"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=pydub"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=moviepy"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=moviepy.editor"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=imageio"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=imageio_ffmpeg"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=decorator"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=proglog"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=urllib.request"
REM Added for single-instance (pywin32)
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=win32event"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=win32api"
set "HIDDEN_IMPORTS=!HIDDEN_IMPORTS! --hidden-import=winerror"
echo Hidden imports set.
echo ---
REM ===================================================================
REM =================== PRE-BUILD VALIDATION ======================
REM ===================================================================
echo.
echo Validating configured paths...
if /I "%COMPILE_WORMHOLE%" == "YES" (
    echo Checking for Wormhole assets...
    REM Removed assets checks since they are not used
    echo No extra assets to validate for Wormhole.
)
REM TODO: Add validation for ASCII and EDITOR assets here
echo Validation complete.
echo.
REM ===================================================================
REM ========================= BUILD PROCESS ===========================
REM ===================================================================
REM 1. Compile Wormhole
if /I "%COMPILE_WORMHOLE%" == "YES" (
    echo.
    echo.
    echo Compiling %WORMHOLE_SCRIPT%...
    echo.
   
    REM Use 'py -m PyInstaller' to run the module
    py -m PyInstaller --noconfirm --onefile --windowed ^
        --icon "%WORMHOLE_ICON%" ^
        --clean ^
        !HIDDEN_IMPORTS! ^
        --distpath "%OUTPUT_DIR%" ^
        --workpath "%LOG_DIR%\build\%WORMHOLE_BUILD_NAME%" ^
        --specpath "%LOG_DIR%" ^
        "%SCRIPT_DIR%\%WORMHOLE_SCRIPT%"
    REM Check for failure
    if %errorlevel% neq 0 (
        echo.
        echo ▩▩▩ ERROR: Failed to compile %WORMHOLE_SCRIPT%. See output above. ▩▩▩
        pause
        goto :eof
    )
    echo Successfully compiled %WORMHOLE_SCRIPT% -> %OUTPUT_DIR%
)
REM If you add them, they will also use the !HIDDEN_IMPORTS! variable.
REM Final
echo.
echo ===============================
echo ✓ Build process finished!
echo Executables: %OUTPUT_DIR%
echo Logs and temporary build files: %LOG_DIR%
echo ===============================
pause
goto :eof