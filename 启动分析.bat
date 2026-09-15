@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
  set "PYTHON=py -3"
) else (
  where python >nul 2>nul
  if errorlevel 1 goto no_python
  set "PYTHON=python"
)

if not exist ".venv\Scripts\python.exe" (
  echo 首次运行，正在创建本地运行环境……
  %PYTHON% -m venv .venv
  if errorlevel 1 goto failed
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto failed
)

echo 正在启动拼多多订单经营分析……
".venv\Scripts\python.exe" app.py
goto end

:no_python
echo 未检测到 Python。请先从 https://www.python.org/downloads/windows/ 安装 Python 3.10 或更高版本。
echo 安装时请勾选“Add Python to PATH”。
pause
exit /b 1

:failed
echo 安装或启动失败，请保留此窗口中的提示信息。
pause
exit /b 1

:end
endlocal
