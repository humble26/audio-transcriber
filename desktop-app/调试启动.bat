@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 调试模式（保留控制台输出，便于看报错）...
python main.py
pause