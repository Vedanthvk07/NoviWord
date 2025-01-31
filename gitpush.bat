@echo off
call npm run build
git add .

:: Prompt the user for a commit message
set /p commit_msg="Enter commit message: "

git commit -m "%commit_msg%"
git push