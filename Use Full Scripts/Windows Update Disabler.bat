@echo off
setlocal EnableDelayedExpansion

:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
    IF "%PROCESSOR_ARCHITECTURE%" EQU "amd64" (
>nul 2>&1 "%SYSTEMROOT%\SysWOW64\cacls.exe" "%SYSTEMROOT%\SysWOW64\config\system"
) ELSE (
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
)

REM --> If error flag set, we do not have admin.

if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params= %*
    echo UAC.ShellExecute "cmd.exe", "/c ""%~s0"" %params:"=""%", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B
	
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:gotAdmin
    pushd "%CD%"
    CD /D "%~dp0"
:-------------------------------------- 

 

echo "----------------------  Windows Updates Disabler  ----------------------"
echo.
echo.
echo "Select an option"
echo.
echo "1) Disable Windows Updates and Automatically Disable it during boot ( To stop windows update from automatically turning on in newer windows 10 builds)."
echo.
echo "2) Disable Windows Updates Once ( No automatic turn on prevention )."
echo.
echo "3) Enable Windows Updates and Let Windows update to turn on automatically."
echo.
echo "4) Enable Windows Updates Once ( Keep automatic turn on prevention )."
echo.
echo.
SET /P _input= "Press option number : "

IF "%_input%" == "1"  (

Rem create turnoffwindowsupdates.bat file
echo "___________________  Disabling Windows Update___________________________" > C:\Scripts\turnoffwindowsupdates.bat
echo. >> C:\Scripts\turnoffwindowsupdates.bat
echo. >> C:\Scripts\turnoffwindowsupdates.bat
echo sc config wuauserv start= disabled >> C:\Scripts\turnoffwindowsupdates.bat
echo net stop wuauserv >> C:\Scripts\turnoffwindowsupdates.bat
echo sc config UsoSvc start= disabled >> C:\Scripts\turnoffwindowsupdates.bat
echo net stop UsoSvc >> C:\Scripts\turnoffwindowsupdates.bat

Rem Add scheduled task entry to run the turnoffwindowsupdates.bat script on every boot

schtasks /create /tn TurnWindowsUpdateOff  /F /sc ONLOGON /DELAY 0001:30 /RL HIGHEST /tr C:\Scripts\turnoffwindowsupdates.bat

schtasks /query /xml /TN TurnWindowsUpdateOff > task.xml

copy task.xml tasktemp.xml

(for /F "delims=" %%a in (tasktemp.xml) do (

   set "line=%%a"
   set "command=!line:Command>=!"
   set "disallowbattery=!line:DisallowStartIfOnBatteries>=!"
   set "stoponbattery=!line:StopIfGoingOnBatteries>=!"
   set "startwhenavailable=<StartWhenAvailable>true</StartWhenAvailable>"
   if "!command!" neq "!line!" (
		set "command=<Command>C:\Scripts\turnoffwindowsupdates.bat</Command>"
		set "anotherLine=<WorkingDirectory>C:\Scripts\</WorkingDirectory>"
		echo !command!
		echo !anotherLine!
   ) else if "!disallowbattery!" neq "!line!" (
		set "disallowbattery=<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
		echo !disallowbattery!
   ) else if "!stoponbattery!" neq "!line!" (
		set "stoponbattery=<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>"
		echo !stoponbattery!
		echo !startwhenavailable!
   ) else (
		echo !command!
   )
   
)) > task.xml

schtasks.exe /Create /XML task.xml /tn TurnWindowsUpdateOff /F

del /F task.xml
del /F tasktemp.xml



Rem Run the script to disable windows updates
call C:\Scripts\turnoffwindowsupdates.bat

echo.
echo "Windows updates disabled and automatic turn on prevention script is added."
echo.

)

IF "%_input%"== "2"  (

Rem Run the script to disable windows updates
sc config wuauserv start= disabled
net stop wuauserv

sc config UsoSvc start= disabled
net stop UsoSvc

echo.
echo "Windows updates disabled."
echo.

)


IF "%_input%"== "3"  (

Rem Remove scheduled task entry to run the turnoffwindowsupdates.bat script on every boot
schtasks /delete /tn TurnWindowsUpdateOff /F

Rem Run the script to enable windows updates
sc config wuauserv start= auto
net start wuauserv

sc config UsoSvc start= auto
net start UsoSvc

echo.
echo "Windows updates enabled and automatic turn on prevention script is removed."
echo.

)



IF "%_input%"== "4"  (

Rem Run the script to enable windows updates
sc config wuauserv start= auto
net start wuauserv

sc config UsoSvc start= auto
net start UsoSvc

echo.
echo "Windows updates enabled."
echo.

)


echo.
echo.
echo.
echo.
echo.

pause



