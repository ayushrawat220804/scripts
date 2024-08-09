@echo off 
Title User Input
echo this script will take users name as input and greet them with a predefined sentence:
pause
cls
:jump
set /p input=Enter the name: 
echo Hey!! %input%, we are happy to have you in the party toonight!!
pause
cls
goto jump 
