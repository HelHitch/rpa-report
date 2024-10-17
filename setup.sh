#!/bin/bash
sudo apt install python3-pip
python3 -m venv venv
python3 -m pip install -r app/requirements.txt
python3 -m PyInstaller --onefile --icon="app/icon.ico" -w --distpath="$(dirname "$(pwd)")" -n="rRPA" --clean app/main.py