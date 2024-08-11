@echo off 
Title Create New Folder
rem simple method 1:
mkdir C:\Users\rohit\OneDrive\Documents\test\test123
pause
rem method 2:
set /p folder_name=Enter new folder name:
set /p folder_path=Enter new folder path:

set new_path=%folder_path%\%folder_name%

mkdir %new_path%
echo %new_path% created
pause 
 
