@echo off
cd /d %~dp0

:: Drop the all changes in the working directory
git reset --hard HEAD

:: Change the remote URL to HTTPS
git remote set-url origin https://github.com/AnnaChen1228/MOST_committee.git

:: Pull the latest changes from the remote repository
git pull

pause
