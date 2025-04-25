#!/bin/bash

if [ ! -d "venv" ]; then
    echo "正在创建虚拟环境"
    python3 -m venv venv
    echo "venv环境创建成功！"
    sleep 2

. ./venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
# python main.py