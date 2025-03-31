@echo off

git pull

py pmojo.py

git add .
git commit -m "Auto-update database"
git push