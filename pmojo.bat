@echo off

git pull

python.exe -m pip install --upgrade pip
pip install keyboard pywinctl ahk requests beautifulsoup4 lxml

py pmojo.py

git add days.db
git commit -m "Auto-update database"
git push
