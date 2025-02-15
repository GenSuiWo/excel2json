# Excel 转 JSON 工具

一个简单的命令行工具，用于将 Excel 文件转换为 JSON 格式。

## 更新日志

### 2025-02-15 v1.1.0

- ✨ 新增多工作表支持：自动将每个工作表转换为独立的 JSON 文件
- 🎨 优化命令行界面：添加加载动画和彩色输出
- 🔧 改进文件命名：输出文件格式为 `原文件名_工作表名.json`
- ⚡️ 性能优化：使用异步文件写入
- 🎯 简化输出：移除数据预览功能，专注于转换进度和结果显示

## 功能特点

- 支持多种 Excel 格式：`.xlsx`, `.xls`, `.xlsb`, `.xlsm`, `.csv`
- 自动检测并处理所有工作表
- 正确处理日期和数字格式
- 支持空值处理
- 跨平台支持（Windows, macOS, Linux）
- 美观的命令行界面，支持加载动画

## 先决条件

1. 安装 [Node.js](https://nodejs.org/) (建议版本 >= 14)
2. 安装 npm (通常随 Node.js 一起安装)

检查安装：
``` bash
node --version
npm --version
```

## 打包方式

### macOS/Linux 系统

1. 克隆或下载项目后，进入项目目录
2. 添加执行权限：
   ```bash
   chmod +x build.sh
   ```
3. 运行打包脚本：
   ```bash
   ./build.sh
   ```
4. 打包完成后，可执行文件位于 `release` 目录：
   - `excel-to-json-macos-x64`: macOS 版本
   - `excel-to-json-linux-x64`: Linux 版本
   - `excel-to-json-win-x64.exe`: Windows 版本

### Windows 系统

1. 克隆或下载项目后，进入项目目录
2. 双击运行 `build.bat`
3. 打包完成后，可执行文件位于 `release` 目录：
   - `excel-to-json-win-x64.exe`: Windows 版本
   - `excel-to-json-macos-x64`: macOS 版本
   - `excel-to-json-linux-x64`: Linux 版本

## 使用方法

1. 运行对应平台的可执行文件
2. 在提示符下输入 Excel 文件路径（可以拖拽文件到窗口）
3. 程序会自动转换并在同目录下生成 JSON 文件
4. 输入 'q' 退出程序

## 注意事项

1. macOS 用户首次运行可能需要在"系统偏好设置"中允许运行
2. Windows 用户如果看到安全提示，需要点击"更多信息"然后"仍要运行"
3. 建议将可执行文件放在单独的目录中
4. 确保对输出目录有写入权限

## 开发相关

如果要修改代码后重新打包：

1. 安装依赖：
   ```bash
   npm install
   ```

2. 安装打包工具：
   ```bash
   npm install -g pkg
   ```

3. 使用打包脚本重新打包

## 支持格式

- `.xlsx`: Excel 2007+ XML 格式
- `.xls`: Excel 2.0-2003 二进制格式
- `.xlsb`: Excel 2007+ 二进制格式
- `.xlsm`: Excel 2007+ 启用宏的工作簿
- `.csv`: 逗号分隔值文件