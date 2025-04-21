@echo off

python3 -m venv venv
venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
python main.py