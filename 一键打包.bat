@echo off
chcp 65001 >nul
set "ROOT=%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 -m venv "%ROOT%.venv"
) else (
    python -m venv "%ROOT%.venv"
)
if errorlevel 1 exit /b 1
"%ROOT%.venv\Scripts\python.exe" -m pip install -r "%ROOT%requirements.txt"
if errorlevel 1 exit /b 1
"%ROOT%.venv\Scripts\python.exe" "%ROOT%build.py"
if errorlevel 1 exit /b 1
echo 构建完成：%ROOT%..\Visio转PDF.exe
pause
