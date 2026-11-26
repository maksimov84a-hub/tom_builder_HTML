@echo off
rem Сборка тома по ГОСТ Р 21.101-2020.
rem Требуется Python 3.12 (см. PYTHON.md): зависимости в vendor/ собраны под cp312.
rem Порядок поиска: %TOM_PYTHON% -> py -3.12 -> типичные пути установки.
rem Программа запускается отдельным процессом без консоли, поэтому окно этого
rem файла закрывается сразу после запуска.
chcp 65001 >nul
setlocal
cd /d "%~dp0"

rem --- Поиск интерпретатора ---
set "PY="
if defined TOM_PYTHON if exist "%TOM_PYTHON%" set "PY=%TOM_PYTHON%"
if not defined PY for /f "delims=" %%P in ('py -3.12 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%P"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY if exist "C:\Python312\python.exe" set "PY=C:\Python312\python.exe"
if not defined PY if exist "C:\Program Files\Python312\python.exe" set "PY=C:\Program Files\Python312\python.exe"
if not defined PY goto :notfound

rem --- pythonw.exe: тот же интерпретатор без консольного окна ---
for %%I in ("%PY%") do set "PYDIR=%%~dpI"
set "PYW=%PYDIR%pythonw.exe"
if not exist "%PYW%" set "PYW="
if defined PYW goto :run_windowless
goto :run_console

:run_windowless
start "" "%PYW%" "%~dp0main.py"
goto :eof

:run_console
rem pythonw.exe не найден — запускаем python.exe в его собственном окне.
start "" "%PY%" "%~dp0main.py"
goto :eof

:notfound
echo.
echo Python 3.12 не найден.
echo Установите его: winget install Python.Python.3.12
echo Другой интерпретатор можно указать переменной TOM_PYTHON.
echo.
pause
exit /b 1
