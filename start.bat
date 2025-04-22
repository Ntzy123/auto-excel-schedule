@echo on

python3 -m venv venv
call venv\Scripts\activate
pip install --upgrade pip
pip install -r requirements.txt
REM python main.py