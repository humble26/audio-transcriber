@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在安装依赖...
python -m pip install -r requirements.txt
echo.
echo 安装完成。按任意键退出。
pause >nul