@echo off
cd /d "%~dp0"
py -3 istemci.py
if errorlevel 1 python istemci.py
if errorlevel 1 pause
