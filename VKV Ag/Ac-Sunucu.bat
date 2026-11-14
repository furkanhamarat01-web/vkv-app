@echo off
cd /d "%~dp0"
py -3 sunucu.py
if errorlevel 1 python sunucu.py
if errorlevel 1 pause
