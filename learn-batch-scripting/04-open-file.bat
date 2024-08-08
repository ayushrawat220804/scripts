@echo off
Title Open File
echo this script will open a pdf and a excel file at once !!
echo first file opens as pdf 
pause
start msedge.exe "C:\Users\rohit\OneDrive\Documents\file encoder decoder\ayushrawatfco.pdf"
echo now this will open will open excel with welcome-to-excel file by default
pause 
start excel.exe "C:\Users\rohit\OneDrive\Documents\welcome-to-excel.xlsx"
exit