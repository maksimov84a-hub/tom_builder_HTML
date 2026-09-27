@echo off
setlocal
title Install DeepSeek Harness shortcut
chcp 65001 >nul

rem  Recreates the "DeepSeek Harness" shortcut on the Desktop and in the
rem  Start Menu, pointing at deepseek-harness.cmd in this folder.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0make-shortcut.ps1" -Destination "%USERPROFILE%\Desktop"
if errorlevel 1 goto :fail
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0make-shortcut.ps1" -Destination "%APPDATA%\Microsoft\Windows\Start Menu\Programs"
if errorlevel 1 goto :fail

echo.
echo Done. A "DeepSeek Harness" shortcut is now on your Desktop and in the Start Menu.
pause
exit /b 0

:fail
echo.
echo [ERROR] Could not create the shortcut.
pause
exit /b 1
