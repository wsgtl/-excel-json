@echo off
chcp 65001 >nul
title 项目创建工具

echo.
echo ===============================================
echo            Excel转JSON项目创建工具
echo ===============================================
echo.

setlocal

:: 检查Node.js是否安装
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 错误：未检测到Node.js
    echo 请先安装Node.js并从官网下载：https://nodejs.org/
    echo.
    pause
    exit /b 1
)

:: 检查工具文件是否存在
if not exist "project-generator.js" (
    echo ❌ 错误：找不到project-generator.js
    echo 请确保批处理文件在工具根目录中
    echo.
    pause
    exit /b 1
)

:: 获取项目名称
set /p PROJECT_NAME=请输入项目名称： 
if "%PROJECT_NAME%"=="" (
    echo ❌ 错误：项目名称不能为空
    echo.
    pause
    exit /b 1
)

:: 创建项目
echo.
echo 🎯 正在创建项目：%PROJECT_NAME%
node project-generator.js new "%PROJECT_NAME%"

if %errorlevel% equ 0 (
    echo.
    echo ✅ 项目创建成功！
    echo 📁 项目路径：projects\%PROJECT_NAME%
    echo.
    echo 📝 使用方法：
    echo    1. 将Excel文件放入 projects\%PROJECT_NAME%\excels 目录
    echo    2. 双击运行 projects\%PROJECT_NAME%\convert.bat
    echo    3. 查看生成的JSON文件在 projects\%PROJECT_NAME%\jsons 目录
) else (
    echo.
    echo ❌ 项目创建失败！
)

echo.
pause