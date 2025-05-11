@echo off
echo === IDM Activation Fix Tool Launcher ===
echo This will launch the fix with proper administrator privileges.
echo.

:: Check for admin rights
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if %errorlevel% neq 0 (
    echo Requesting administrative privileges...
    echo Please click "Yes" when prompted to allow admin access.
    powershell -Command "Start-Process -FilePath '%~0' -Verb RunAs"
    exit /b
)

:: If we get here, we have admin rights
echo Running IDM Fix with administrator privileges...
echo.

:: Run the Python script
python "%~dp0IDM-Fix.py"

pause 