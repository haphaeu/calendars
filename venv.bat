@echo off

cd /d %~dp0

choice /c yn /t 5 /d y /m "Open Jupyter?"
if ERRORLEVEL 2 goto END

call .\venv\scripts\activate.bat
start "" jupyter notebook

:END

start "" .\venv\scripts\activate.bat
