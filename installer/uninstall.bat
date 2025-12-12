@echo off
REM ==============================================================================
REM Grammar & QS Checker Add-in Uninstaller (Windows)
REM ==============================================================================

SETLOCAL

echo.
echo ==========================================
echo  Grammar ^& QS Checker Add-in Uninstaller
echo ==========================================
echo.

SET ADDIN_NAME=GrammarChecker_QS.xlam
SET INSTALL_DIR=%APPDATA%\Microsoft\AddIns

echo This will remove the Grammar ^& QS Checker add-in from:
echo %INSTALL_DIR%\%ADDIN_NAME%
echo.

choice /C YN /M "Do you want to continue?"
IF ERRORLEVEL 2 (
    echo Uninstallation cancelled.
    pause
    exit /b 0
)

echo.
echo Removing add-in...

IF EXIST "%INSTALL_DIR%\%ADDIN_NAME%" (
    del "%INSTALL_DIR%\%ADDIN_NAME%"
    echo Add-in removed successfully!
) ELSE (
    echo Add-in file not found. It may have been already removed.
)

echo.
echo ==========================================
echo  Uninstallation Complete
echo ==========================================
echo.
echo The add-in has been removed from your system.
echo.
echo If Excel is running, please restart it.
echo.
echo To remove the add-in from Excel's list:
echo 1. Open Excel
echo 2. Go to: File ^> Options ^> Add-ins
echo 3. At the bottom, select "Excel Add-ins" and click "Go..."
echo 4. Uncheck "GrammarChecker_QS" if it appears
echo 5. Click OK
echo.

pause
ENDLOCAL
