@echo off
chcp 65001 >nul 2>&1
if not exist venv (
    echo "正在创建虚拟环境"
    python3 -m venv venv
    echo "venv环境创建成功！"
    timeout /t 2 >nul
)
call venv\Scripts\activate
echo "正在检查并安装依赖包"
pip install --upgrade pip
pip install -r requirements.txt

REM 选择打包或退出
cls
echo ============================
echo 1. 打包为可执行文件
echo 2. 退出
echo ============================
choice /c 12 /n /m "请选择 [1 or 2]: "

if errorlevel 2 goto exit
if errorlevel 1 goto run_command

:run_command
pyinstaller --onefile --name=auto-excel-schedule main.py
echo "打包完成，请按任意键继续..."
pause >nul
exit

:exit
echo "请按任意键继续..."
pause >nul
exit

REM python main.py