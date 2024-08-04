@echo off
setlocal enabledelayedexpansion

REM ================================
REM Check for Administrative Privileges
REM ================================
REM Check if the script is running with administrative privileges
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo This script requires administrative privileges. Please run it as an administrator.
    pause
    exit /b
)


REM ================================
REM Title and Introduction
REM ================================
echo.
echo ================================
echo        Developed By Muhammad
echo ================================
echo.
echo ================================
echo        Batch File Execution
echo ================================
echo.
echo This script will:
echo 1. Activate Windows
echo 2. Download Office Deployment Tool (ODT)
echo 3. Prompt you to run ODT manually
echo 4. Move generated files and create a new configuration file
echo 5. Run Office setup with the new configuration
echo.

REM ================================
REM Step 1: Activate Windows
REM ================================
echo.
echo ================================
echo        Windows Activation
echo ================================
echo.
SET /P ACTIVATE_WINDOWS=Do you want to activate Windows? (Y/[N]): 
SET ACTIVATE_WINDOWS=%ACTIVATE_WINDOWS: =%
SET ACTIVATE_WINDOWS=%ACTIVATE_WINDOWS:~0,1%
IF /I "%ACTIVATE_WINDOWS%" EQU "Y" (
    echo Activating Windows...
    slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
    IF ERRORLEVEL 1 (
        echo Failed to set the product key. Exiting...
        GOTO END
    )
    slmgr /skms kms8.msguides.com
    IF ERRORLEVEL 1 (
        echo Failed to set the KMS server. Exiting...
        GOTO END
    )
    slmgr /ato
    IF ERRORLEVEL 1 (
        echo Activation failed. Exiting...
        GOTO END
    )
    echo Windows has been activated.
) ELSE (
    echo Windows activation skipped.
)

REM ================================
REM Step 2: Prompt User to Install Office
REM ================================
echo.
echo ================================
echo      Office Installation Prompt
echo ================================
echo.
SET /P INSTALL_OFFICE=Do you want to install Office 2024 now? (Y/[N]): 
SET INSTALL_OFFICE=%INSTALL_OFFICE: =%
SET INSTALL_OFFICE=%INSTALL_OFFICE:~0,1%
IF /I "%INSTALL_OFFICE%" NEQ "Y" (
    echo Office installation skipped.
    GOTO END
)

REM ================================
REM Step 3: Download Office Deployment Tool (ODT)
REM ================================
echo.
echo ================================
echo    Download Office Deployment Tool
echo ================================
echo.
echo Downloading Office Deployment Tool (ODT)...
powershell -command "Invoke-WebRequest -Uri 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe' -OutFile '%USERPROFILE%\Downloads\officedeploymenttool.exe'"
IF ERRORLEVEL 1 (
    echo Failed to download Office Deployment Tool. Exiting...
    GOTO END
)

REM ================================
REM Step 4: Prompt User to Run ODT Manually
REM ================================
echo.
echo ================================
echo  Run Office Deployment Tool Manually
echo ================================
echo.
echo Please run the Office Deployment Tool manually from the Downloads folder to generate configuration files.
echo After running the ODT, press any key to continue...
pause

REM ================================
REM Step 5: Move Generated Files to C:\Office
REM ================================
echo.
echo ================================
echo    Move Files to C:\Office
echo ================================
echo.
echo Moving generated files to C:\Office...
mkdir C:\Office
xcopy "%USERPROFILE%\Downloads\*" "C:\Office\" /S /I /Y
IF ERRORLEVEL 1 (
    echo Failed to move files to C:\Office. Exiting...
    GOTO END
)

REM ================================
REM Step 6: Remove Unnecessary Configuration Files
REM ================================
echo.
echo ================================
echo   Remove Unnecessary Configuration Files
echo ================================
echo.
echo Removing unnecessary configuration files...
del "C:\Office\configuration-Office2021Enterprise.xml" >nul 2>&1
del "C:\Office\configuration-Office2019Enterprise.xml" >nul 2>&1
del "C:\Office\configuration-Office365-x86.xml" >nul 2>&1

REM ================================
REM Step 7: Create New Configuration File
REM ================================
echo.
echo ================================
echo      Create New Configuration File
echo ================================
echo.
echo Creating new configuration file...
(
    echo ^<Configuration^>
    echo   ^<Add OfficeClientEdition="64" Channel="PerpetualVL2024"^>
    echo     ^<Product ID="ProPlus2024Volume" PIDKEY="2TDPW-NDQ7G-FMG99-DXQ7M-TX3T2"^>
    echo       ^<Language ID="en-us"/^>
    echo     ^</Product^>
    echo   ^</Add^>
    echo   ^<RemoveMSI /^>
    echo   ^<Property Name="AUTOACTIVATE" Value="1" /^>
    echo ^</Configuration^>
) > "C:\Office\office-setup.xml"

REM ================================
REM Step 8: Run setup.exe with the New Configuration
REM ================================
echo.
echo ================================
echo     Run Office Installer with Config
echo ================================
echo.
echo Running setup.exe with configuration. Kindly dont close the command prompt untill office is installed and the next prompt arrives.
cd /d C:\Office
start /wait setup.exe /configure office-setup.xml
IF ERRORLEVEL 1 (
    echo Failed to install Office. Exiting...
    GOTO END
)
echo Office installation process completed.

REM ================================
REM Step 9: Prompt for Cleanup
REM ================================
echo.
echo ================================
echo     Cleanup Prompt
echo ================================
echo.
SET /P CLEANUP=Do you want the script to remove unnecessary files? (Y/[N]): 
SET CLEANUP=%CLEANUP: =%
SET CLEANUP=%CLEANUP:~0,1%
IF /I "%CLEANUP%" EQU "Y" (
    echo.
    echo ================================
    echo        Cleanup
    echo ================================
    echo.
    echo Removing unnecessary files...
    
    REM Example cleanup commands
    del "%USERPROFILE%\Downloads\officedeploymenttool.exe" >nul 2>&1
    del "%USERPROFILE%\Downloads\configuration-Office365-x64" >nul 2>&1
    del "%USERPROFILE%\Downloads\configuration-Office365-x86.xaml" >nul 2>&1
    del "%USERPROFILE%\Downloads\configuration-Office2019Enterprise.xaml" >nul 2>&1
    del "%USERPROFILE%\Downloads\configuration-Office2021Enterprise.xaml" >nul 2>&1
    rmdir /s /q C:\Office

    echo Cleanup completed.
)

:END
endlocal

REM Pause to keep the command prompt from closing
echo.
echo ================================
echo               Done
echo ================================
echo.
pause