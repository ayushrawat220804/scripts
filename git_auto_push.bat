@echo off
setlocal enabledelayedexpansion

echo Git Auto Push Script
echo This script automates git add, commit, and push operations
echo.

:ask_repo_path
set /p "repo_path=Enter the path to your Git repository: "

:: Remove quotes if present
set repo_path=%repo_path:"=%

:: Check if directory exists
if not exist "%repo_path%" (
    echo Error: Directory does not exist: %repo_path%
    goto ask_repo_path
)

:: Navigate to repository
cd /d "%repo_path%"

:: Check if it's a git repository
git rev-parse --is-inside-work-tree >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Not a valid Git repository: %repo_path%
    goto ask_repo_path
)

:: Check if there are changes to commit
git status --porcelain >nul 2>&1
if %errorlevel% neq 0 (
    echo No changes to commit in %repo_path%
    goto end
)

:: Add all changes
echo Adding all changes...
git add .

:: Generate basic commit message
echo Generating commit message...
set "file_types="
for /f "tokens=*" %%a in ('git diff --cached --name-only') do (
    set "file=%%a"
    set "ext=%%~xa"
    if defined ext (
        set "file_types=!file_types! !ext!"
    )
)

set "commit_message=Update code with changes"

:: Get stats
for /f "tokens=*" %%a in ('git diff --cached --stat ^| find "changed"') do (
    set "stats=%%a"
    set "commit_message=!commit_message! (!stats!)"
)

:: Show commit message and allow user to edit
echo.
echo Generated commit message: %commit_message%
set /p "use_message=Use this message? (Y/n): "

if /i "%use_message%"=="n" (
    set /p "commit_message=Enter your commit message: "
)

:: Commit changes
echo.
echo Committing changes with message: %commit_message%
git commit -m "%commit_message%"

:: Try to push to remote
echo.
echo Pushing changes to remote repository...
git push

:: Check if push failed
if %errorlevel% neq 0 (
    echo.
    echo Push failed. Remote repository may have changes you don't have locally.
    set /p "pull_first=Do you want to pull changes first and then try pushing again? (Y/n): "
    
    if /i not "%pull_first%"=="n" (
        echo.
        echo Pulling changes from remote repository...
        
        :: Try to pull with rebase to avoid merge commit
        git pull --rebase
        
        if %errorlevel% neq 0 (
            echo.
            echo Pull with rebase failed. Trying normal pull...
            git pull
            
            if %errorlevel% neq 0 (
                echo.
                echo Pull failed. You may need to resolve conflicts manually.
                goto end
            )
        )
        
        echo.
        echo Pushing changes to remote repository after pull...
        git push
        
        if %errorlevel% neq 0 (
            echo.
            echo Push failed again. You may need to resolve this manually.
        else
            echo.
            echo Changes successfully pulled and pushed!
        )
    ) else (
        echo.
        echo Push operation was skipped. Changes are committed but not pushed.
    )
) else (
    echo.
    echo Changes successfully committed and pushed!
)

:end
endlocal
pause