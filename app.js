#!/usr/bin/env node
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

// 创建命令行交互接口
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Excel转JSON函数
function excelToJson(excelPath) {
    try {
        console.log(`\n正在读取Excel文件: ${excelPath}`);
        
        // 读取Excel文件，添加一些选项以提高兼容性
        const workbook = xlsx.readFile(excelPath, {
            type: 'binary',
            cellDates: true,  // 正确处理日期
            cellNF: false,    // 不保留数字格式
            cellText: false   // 不保留文本格式
        });

        // 获取所有工作表名称
        const sheetNames = workbook.SheetNames;
        console.log(`\n发现 ${sheetNames.length} 个工作表: ${sheetNames.join(', ')}`);
        
        // 如果有多个工作表，让用户选择
        let sheetName = sheetNames[0];
        if (sheetNames.length > 1) {
            console.log('默认使用第一个工作表，如需其他工作表请重新运行程序');
        }
        
        const worksheet = workbook.Sheets[sheetName];
        
        // 转换为JSON，添加选项以处理空值
        console.log('\n正在转换为JSON格式...');
        const jsonData = xlsx.utils.sheet_to_json(worksheet, {
            defval: null,     // 空单元格的默认值
            raw: false,       // 转换数字和日期
            dateNF: 'YYYY-MM-DD'  // 日期格式
        });
        
        console.log(`转换完成! 共 ${jsonData.length} 行数据`);
        
        // 生成输出路径
        const outputPath = path.join(
            path.dirname(excelPath),
            path.basename(excelPath, path.extname(excelPath)) + '.json'
        );
        
        // 保存JSON文件
        console.log(`\n正在保存到: ${outputPath}`);
        fs.writeFileSync(
            outputPath,
            JSON.stringify(jsonData, null, 2),
            'utf8'
        );
        console.log('保存成功!');
        
        // 显示数据预览
        console.log('\n数据预览 (前3行):');
        console.log(JSON.stringify(jsonData.slice(0, 3), null, 2));
        
        return { success: true, outputPath };
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// 主函数
async function main() {
    console.log('Excel转JSON工具');
    console.log('支持的格式: .xlsx, .xls, .xlsb, .xlsm, .csv');
    console.log('=' .repeat(50));
    
    const promptUser = () => {
        rl.question('\n请输入Excel文件路径 (输入q退出): ', (excelPath) => {
            // 移除输入路径两端的空格和引号
            excelPath = excelPath.trim().replace(/^["']|["']$/g, '');
            
            // 检查是否退出
            if (excelPath.toLowerCase() === 'q') {
                console.log('\n感谢使用，再见!');
                rl.close();
                return;
            }
            
            // 检查文件是否存在
            if (!fs.existsSync(excelPath)) {
                console.log(`\n错误: 找不到文件 '${excelPath}'`);
                promptUser();
                return;
            }
            
            // 检查文件扩展名
            const ext = path.extname(excelPath).toLowerCase();
            const supportedExts = ['.xlsx', '.xls', '.xlsb', '.xlsm', '.csv'];
            if (!supportedExts.includes(ext)) {
                console.log(`\n错误: 不支持的文件格式。支持的格式: ${supportedExts.join(', ')}`);
                promptUser();
                return;
            }
            
            // 执行转换
            const result = excelToJson(excelPath);
            
            if (result.success) {
                console.log(`\n✨ 转换成功! JSON文件已保存至:\n${result.outputPath}`);
            } else {
                console.log(`\n❌ 错误: ${result.error}`);
            }
            
            console.log('\n' + '='.repeat(50));
            promptUser();
        });
    };
    
    promptUser();
}

// 处理Ctrl+C
process.on('SIGINT', () => {
    console.log('\n\n已取消操作，感谢使用!');
    rl.close();
});

// 启动程序
main().catch(error => {
    console.error('程序出错:', error);
    rl.close();
});
