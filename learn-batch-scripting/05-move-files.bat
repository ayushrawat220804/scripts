@echo off
Title Move Files
echo this script will move a single file 
pause 
move "C:\Users\rohit\OneDrive\Documents\text - Copy.txt" C:\Users\rohit\OneDrive\Documents\test
echo this script will move a set of multiple files ending with "-copy" name in it . . .
pause
move "C:\Users\rohit\OneDrive\Documents\test\*- Copy*" C:\Users\rohit\OneDrive\Documents\test2

rem The * (asterisk) is a wildcard character used in batch scripting to represent any number of characters. When you use *-copy*, it matches any file name that contains the string -copy     rem anywhere in the name. Here’s a breakdown:



rem NOTE :
rem *: Matches zero or more characters.
rem-copy: The specific string you’re looking for in the file names.
rem *: Again, matches zero or more characters.
rem So, *-copy* will match file names like:

:: file-copy.txt
:: document-copy.docx
:: image-copy.jpg