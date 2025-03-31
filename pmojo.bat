@echo off

git pull

python.exe -m pip install --upgrade pip
pip install keyboard requests bs4 pywinauto ahk
py pmojo.py

git pull
git add .
git commit -m "Auto-update database"
git push