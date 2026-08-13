@echo off
title DNF Buff动画打包工具
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo    DNF Buff动画随机切换工具 - 一键打包
echo ============================================
echo.

:: ---- 第一步：找到 Python ----
echo [1/5] 检测 Python 环境...

:: 尝试 where python
set "PYTHON_CMD=python"
for /f "delims=" %%i in ('where python 2^>nul') do set "PYTHON_CMD=%%i" & goto :found_py
:: 尝试 where py
for /f "delims=" %%i in ('where py 2^>nul') do set "PYTHON_CMD=%%i -3" & goto :found_py
:: 尝试常见安装路径
if exist "C:\Program Files\Python311\python.exe" set "PYTHON_CMD=C:\Program Files\Python311\python.exe" & goto :found_py
if exist "C:\Program Files\Python310\python.exe" set "PYTHON_CMD=C:\Program Files\Python310\python.exe" & goto :found_py
if exist "C:\Program Files\Python39\python.exe" set "PYTHON_CMD=C:\Program Files\Python39\python.exe" & goto :found_py
if exist "C:\Program Files\Python38\python.exe" set "PYTHON_CMD=C:\Program Files\Python38\python.exe" & goto :found_py
:: 尝试用户目录安装
for /f "delims=" %%i in ('dir /s /b "%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe" 2^>nul') do set "PYTHON_CMD=%%i" & goto :found_py
for /f "delims=" %%i in ('dir /s /b "%USERPROFILE%\AppData\Local\Programs\Python\Python310\python.exe" 2^>nul') do set "PYTHON_CMD=%%i" & goto :found_py
:: 尝试注册表
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Python\PythonCore\3.11\InstallPath" /ve 2^>nul') do if exist "%%bpython.exe" set "PYTHON_CMD=%%bpython.exe" & goto :found_py
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Python\PythonCore\3.10\InstallPath" /ve 2^>nul') do if exist "%%bpython.exe" set "PYTHON_CMD=%%bpython.exe" & goto :found_py

echo [错误] 未找到 Python，请先安装 Python 3.8-3.11 并勾选"Add to PATH"
echo        安装后重新双击本脚本即可。
echo.
echo        下载地址: https://www.python.org/downloads/
pause
exit /b 1
:found_py
echo         找到 Python: %PYTHON_CMD%

:: ---- 第二步：安装依赖 ----
echo [2/5] 检查并安装依赖库...
%PYTHON_CMD% -c "import PyQt5, openpyxl, imageio_ffmpeg" >nul 2>nul
if errorlevel 1 (
    echo         正在安装依赖库，请稍候...
    %PYTHON_CMD% -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖库安装失败
        pause
        exit /b 1
    )
    echo         依赖库安装完成
) else (
    echo         依赖库已就绪
)

:: ---- 第三步：安装 PyInstaller ----
echo [3/5] 检查 PyInstaller...
%PYTHON_CMD% -m PyInstaller --version >nul 2>nul
if errorlevel 1 (
    echo         正在安装 PyInstaller...
    %PYTHON_CMD% -m pip install pyinstaller
    if errorlevel 1 (
        echo [错误] PyInstaller 安装失败
        pause
        exit /b 1
    )
) else (
    echo         PyInstaller 已就绪
)

:: ---- 第四步：清理旧构建 ----
echo [4/5] 清理旧构建产物...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

:: ---- 第五步：执行打包 ----
echo [5/5] 开始打包，请耐心等待（约 1-3 分钟）...
echo.
%PYTHON_CMD% -m PyInstaller DNFBuffSwitcher.spec --noconfirm --clean
if errorlevel 1 (
    echo [错误] 打包失败，请检查上方日志
    pause
    exit /b 1
)

:: 复制对应表到输出目录
if exist "BUFF动画职业名对照表.xlsx" (
    copy /y "BUFF动画职业名对照表.xlsx" "dist\" >nul
    echo         对应表已复制
)

echo.
echo ============================================
echo    打包成功!
echo.
echo    主程序: dist\DNFBuffSwitcher.exe
echo    对应表: dist\BUFF动画职业名对照表.xlsx
echo ============================================
echo.
echo 按任意键退出...
pause >nul
exit /b 0