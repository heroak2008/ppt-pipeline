@echo off
rem ppt-lite 一键启动（Windows）
rem 用法：双击或命令行运行 start.bat；可选端口参数 start.bat 9000
setlocal
cd /d "%~dp0"

where uv >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 uv。请先安装：https://docs.astral.sh/uv/
    pause
    exit /b 1
)

set PORT=%1
if "%PORT%"=="" set PORT=8765

echo [ppt-lite] 同步依赖（首次较慢，使用清华镜像）...
uv sync
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络。
    pause
    exit /b 1
)

echo [ppt-lite] 启动服务：http://127.0.0.1:%PORT%
uv run uvicorn ppt_lite.app:app --host 127.0.0.1 --port %PORT% --workers 1
pause
