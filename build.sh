#!/bin/bash

# 安装依赖
npm install

# 安装pkg（如果没有安装）
if ! command -v pkg &> /dev/null; then
    echo "正在安装pkg..."
    sudo npm install -g pkg
fi

# 执行打包
echo "开始打包..."
npm run build

# 创建发布目录
mkdir -p release

# 移动生成的文件到release目录
mv excel-to-json-* release/

echo "打包完成！文件在release目录中："
ls -l release/

echo "
Windows版本: release/excel-to-json-win-x64.exe
MacOS版本: release/excel-to-json-macos-x64
Linux版本: release/excel-to-json-linux-x64
" 