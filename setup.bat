@echo off
REM 安装pyinstaller
pip install pyinstaller

REM 打包应用程序
pyinstaller --onefile --windowed --icon=app_icon.ico app.py

REM 暂停以查看输出
pause