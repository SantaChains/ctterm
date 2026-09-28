@echo off
rem EldenCT GUI 启动器（由 explorer 拉起可脱离父会话存活）
rem %~dp0 = 本 bat 所在目录，与放置位置解耦
cd /d "%~dp0"
start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0_gui_launch.pyw"
