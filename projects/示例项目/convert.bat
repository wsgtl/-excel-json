@echo off
chcp 65001 >nul
title Excel转JSON转换工具

echo.
echo ===============================================
echo            Excel 转 JSON 转换工具
echo ===============================================
echo.

setlocal

:: 设置颜色
set "color_header=0E"
set "color_success=0A"
set "color_error=0C"
set "color_warning=0E"
set "color_info=0B"

:: 获取当前目录和工具根目录
set "CURRENT_DIR=%~dp0"
set "TOOL_ROOT=%CURRENT_DIR%..\.."

:: 检查工具根目录是否存在
if not exist "%TOOL_ROOT%\excel2json.js" (
    echo.
    call :color_echo "❌ 错误：找不到转换工具" %color_error%
    echo    请确保批处理文件在正确的项目目录中
    echo    目录结构应该是：projects/你的项目/convert.bat
    echo.
    pause
    exit /b 1
)

:: 定义目录
set "EXCEL_DIR=%CURRENT_DIR%excels"
set "JSON_DIR=%CURRENT_DIR%jsons"

:: 检查Excel目录是否存在
if not exist "%EXCEL_DIR%" (
    call :color_echo "📁 创建Excel目录..." %color_info%
    mkdir "%EXCEL_DIR%"
    call :color_echo "✅ 已创建Excel目录：%EXCEL_DIR%" %color_success%
    call :color_echo "💡 请将Excel文件放入此目录，然后重新运行此脚本" %color_warning%
    echo.
    pause
    exit /b 0
)

:: 创建JSON输出目录
if not exist "%JSON_DIR%" (
    mkdir "%JSON_DIR%"
)

call :color_echo "🔍 检查Excel文件..." %color_info%

:: 查找Excel文件
dir /b "%EXCEL_DIR%\*.xlsx" "%EXCEL_DIR%\*.xls" >nul 2>&1
if %errorlevel% neq 0 (
    call :color_echo "❌ 在Excel目录中未找到任何.xlsx或.xls文件" %color_error%
    echo    请将Excel文件放入：%EXCEL_DIR%
    echo.
    pause
    exit /b 1
)

:: 显示找到的Excel文件
call :color_echo "📊 找到以下Excel文件：" %color_success%
for %%f in ("%EXCEL_DIR%\*.xlsx" "%EXCEL_DIR%\*.xls") do (
    echo     📄 %%~nxf
)

echo.
call :color_echo "🔄 开始转换过程..." %color_header%
echo.

:: 执行转换
node "%TOOL_ROOT%\excel2json.js" convert -i "%EXCEL_DIR%" -o "%JSON_DIR%"

set "CONVERT_RESULT=%errorlevel%"

echo.
if %CONVERT_RESULT% equ 0 (
    call :color_echo "✅ 转换完成！" %color_success%
    echo.
    call :color_echo "📁 生成的JSON文件位置：" %color_info%
    echo     %JSON_DIR%
    echo.
    
    :: 显示生成的JSON文件
    if exist "%JSON_DIR%\*.json" (
        call :color_echo "📄 生成的JSON文件：" %color_success%
        for %%j in ("%JSON_DIR%\*.json") do (
            echo     📋 %%~nxj
        )
    )
) else (
    call :color_echo "❌ 转换过程中出现错误" %color_error%
)

echo.
call :color_echo "⏰ 完成时间: %date% %time%" %color_info%
echo.

pause
exit /b %CONVERT_RESULT%

:: 颜色输出函数
:color_echo
set "msg=%~1"
set "color=%~2"
echo %msg%
exit /b 0

endlocal