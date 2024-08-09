@echo off 
Title Yes or No
echo this script will ask user yes or no!!
pause 
:start
cls
set /p user_input=Do you want to continue (y/n)?:  
if not defined user_input goto start
if /i %user_input%==y goto yes
if /i %user_input%==n (goto no) else (goto invalid)

:yes
echo %user_input% user have entered y. 
pause
goto start

:no
echo %user_input% user have entered n.
pause 
exit 

:invalid
echo %user_input% is invalid input, try again!!
set user_input= 
pause 
goto start


rem NOTE: 
rem /i means you can write both lower and uppercase eg y or Y.
rem /p means prompt