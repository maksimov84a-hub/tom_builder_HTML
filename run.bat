@echo off
cd /d "%~dp0"
set "PYTHON=C:\Users\Aleksandr\AppData\Local\Programs\Python\Python312\python.exe"
if exist "%PYTHON%" (
    "%PYTHON%" main.py
) else (
    python main.py
)
