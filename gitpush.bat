@echo off
call npm run build
git add .
git commit -m "updated the file"
git push