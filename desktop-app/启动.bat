@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动「音频转写」桌面端...
start "" pythonw main.py
if errorlevel 1 (
  echo 启动失败，改用带控制台方式运行：
  python main.py
  pause
)