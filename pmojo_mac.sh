#!/bin/bash
set -e

cd "$(dirname "$0")"

git pull

# Activate venv if it exists, otherwise use system python
if [ -d "venv14" ]; then
    source venv14/bin/activate
fi

pip install --upgrade pip
pip install pywinctl pyautogui requests beautifulsoup4 lxml

python pmojo.py

git add days.db
git commit -m "Auto-update database"
git push