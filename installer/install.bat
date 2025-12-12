@echo off
REM ==============================================================================
REM Grammar & QS Checker Add-in Installer (Windows)
REM ==============================================================================

SETLOCAL

echo.
echo ========================================
echo  Grammar ^& QS Checker Add-in Installer
echo ========================================
echo.

REM Check if running as administrator (optional but recommended)
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges...
) else (
    echo Warning: Not running as administrator. Some features may not work.
    echo.
)

REM Define variables
SET ADDIN_NAME=GrammarChecker_QS.xlam
SET ADDIN_PATH=%~dp0%ADDIN_NAME%
SET INSTALL_DIR=%APPDATA%\Microsoft\AddIns

echo Add-in file: %ADDIN_NAME%
echo Install location: %INSTALL_DIR%
echo.

REM Check if add-in file exists
IF NOT EXIST "%ADDIN_PATH%" (
    echo ERROR: Add-in file not found: %ADDIN_PATH%
    echo Please ensure %ADDIN_NAME% is in the same folder as this installer.
    echo.
    pause
    exit /b 1
)

REM Create installation directory if it doesn't exist
IF NOT EXIST "%INSTALL_DIR%" (
    echo Creating installation directory...
    mkdir "%INSTALL_DIR%"
)

REM Check if add-in already installed
IF EXIST "%INSTALL_DIR%\%ADDIN_NAME%" (
    echo.
    echo WARNING: An existing version of the add-in was found.
    choice /C YN /M "Do you want to replace it?"
    IF ERRORLEVEL 2 (
        echo Installation cancelled.
        pause
        exit /b 0
    )
    echo Removing old version...
    del "%INSTALL_DIR%\%ADDIN_NAME%"
)

REM Copy add-in to installation directory
echo.
echo Installing add-in...
copy "%ADDIN_PATH%" "%INSTALL_DIR%\" >nul

IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to copy add-in file.
    echo Please check permissions and try again.
    echo.
    pause
    exit /b 1
)

echo Add-in installed successfully!
echo.
echo ========================================
echo  Installation Complete
echo ========================================
echo.
echo The Grammar ^& QS Checker add-in has been installed to:
echo %INSTALL_DIR%\%ADDIN_NAME%
echo.
echo NEXT STEPS:
echo 1. Restart Microsoft Excel if it's currently running
echo 2. In Excel, go to: File ^> Options ^> Add-ins
echo 3. At the bottom, select "Excel Add-ins" and click "Go..."
echo 4. Check the box next to "GrammarChecker_QS"
echo 5. Click OK
echo.
echo The add-in buttons will appear in the Excel ribbon.
echo.
echo If you encounter security warnings:
echo - Go to File ^> Options ^> Trust Center ^> Trust Center Settings
echo - Select "Macro Settings" and enable macros
echo - Or add %INSTALL_DIR% to Trusted Locations
echo.
echo For help, see the User Guide documentation.
echo.

pause
ENDLOCAL
