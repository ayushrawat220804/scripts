@echo off

if not exist "C:\Program Files\7-Zip" (
echo 7-zip not installed!
pause
exit
)

set /p archive="Enter archive: "
if not exist "%archive%" (
echo archive not found!
pause
exit
)

set /p wordlist="Enter Wordlist: "
if not exist "%archive%" (
echo wordlist not found!
pause
exit
)

for /f %%a in (%wordlist%) do (
set pass = %%a
call :attempr
)
echo shitty wordlist dumbass
pause 
exit

:attempt 
C:\Program Files\7-Zip\7zip.exe" x -p%pass% "%archive% -o"cracked" -y >nul 2>&1
echo cracking...
if /I %errorlevel% EQU 0 (
echo password found: %pass%
pause
ext
)