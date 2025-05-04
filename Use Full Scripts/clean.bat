@echo off
cls
echo Cleaning up Windows... Please wait.

:: Run Disk Cleanup
cleanmgr /sagerun:1

:: Clear Temporary Files
del /s /q %temp%\* > nul 2>&1
del /s /q C:\Windows\Temp\* > nul 2>&1
del /s /q C:\Users\%username%\AppData\Local\Temp\* > nul 2>&1

echo Temp files deleted.

:: Check and Repair System Files
echo Running SFC scan...
sfc /scannow

echo Running DISM...
DISM /Online /Cleanup-Image /RestoreHealth

:: Flush DNS Cache
echo Flushing DNS cache...
ipconfig /flushdns

:: Clear Windows Update Cache
echo Stopping Windows Update services...
net stop wuauserv > nul 2>&1
net stop bits > nul 2>&1
rd /s /q C:\Windows\SoftwareDistribution > nul 2>&1
rd /s /q C:\Windows\System32\catroot2 > nul 2>&1
echo Restarting Windows Update services...
net start wuauserv > nul 2>&1
net start bits > nul 2>&1

echo Windows Update cache cleared.

:: Clear Event Logs
echo Clearing Event Logs...
for /F "tokens=*" %%1 in ('wevtutil el') DO wevtutil cl "%%1"

echo Event logs cleared.

:: Optimize and Defrag System Drive
echo Optimizing system drive...
defrag C: /O

echo Running silent Disk Cleanup...
cleanmgr /verylowdisk

echo Windows cleanup completed successfully!
pause
