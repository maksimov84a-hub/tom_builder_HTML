@echo off
setlocal
title DeepSeek Harness
chcp 65001 >nul

rem  DeepSeek Harness - desktop launcher
rem  Official site: https://www.deepseek.com/harness/en/
rem  Starts the local browser UI with the documented npx invocation.
rem
rem  The CURRENT FOLDER becomes the workspace the agent may read and edit.
rem  The shortcut sets it via Properties > "Start in" (default: this repo root).

echo Starting DeepSeek Harness  ^(npx @deepseek-ai/dsh web^)
echo Workspace: %CD%
echo.

call npx --yes @deepseek-ai/dsh web %*
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" goto :stopped

echo [ERROR] dsh web exited with code %EXITCODE%.
echo.
echo If the port is already in use, another DeepSeek Harness server is most
echo likely already running (the default UI is http://127.0.0.1:3080).
echo Either use that window, or start this one on another port:
echo     deepseek-harness.cmd --port 8080
goto :hold

:stopped
echo dsh web stopped.

:hold
echo Keep this window open while you use the web UI; closing it stops the server.
pause
