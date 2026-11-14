@echo off
cd /d "%~dp0"
where pyw >nul 2>&1 && start "" pyw -3 "%~dp0is_merkezi.py" && exit /b 0
where pythonw >nul 2>&1 && start "" pythonw "%~dp0is_merkezi.py" && exit /b 0
where py >nul 2>&1 && start "" py -3 "%~dp0is_merkezi.py" && exit /b 0
where python >nul 2>&1 && start "" python "%~dp0is_merkezi.py" && exit /b 0
echo Python 3 bulunamadi.
echo python.org/downloads adresinden Python 3 kurun.
echo Kurulumda Add python.exe to PATH secili olsun.
pause
