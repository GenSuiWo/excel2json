@echo off
echo 正在安装依赖...
call npm install

echo 检查pkg是否安装...
where pkg >nul 2>nul
if %errorlevel% neq 0 (
    echo 正在安装pkg...
    call npm install -g pkg
)

echo 开始打包...
call npm run build

echo 创建release目录...
if not exist release mkdir release

echo 移动文件到release目录...
move excel-to-json-* release\

echo 打包完成！文件在release目录中：
dir release

echo.
echo Windows版本: release\excel-to-json-win-x64.exe
echo MacOS版本: release\excel-to-json-macos-x64
echo Linux版本: release\excel-to-json-linux-x64
echo.

pause 