#!/usr/bin/env node
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');
const ora = require('ora');
const chalk = require('chalk');
const axios = require('axios');
const os = require('os');

// 创建命令行交互接口
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Excel转JSON函数
async function excelToJson(excelPath) {
    try {
        // 创建加载动画
        const spinner = ora({
            text: chalk.blue('正在读取Excel文件...'),
            spinner: 'dots'
        }).start();
        
        // 读取Excel文件
        const workbook = xlsx.readFile(excelPath, {
            type: 'binary',
            cellDates: true,
            cellNF: false,
            cellText: false
        });

        // 获取所有工作表名称
        const sheetNames = workbook.SheetNames;
        spinner.succeed(chalk.green(`发现 ${chalk.bold(sheetNames.length)} 个工作表: ${chalk.cyan(sheetNames.join(', '))}`));
        
        const results = [];
        
        // 处理每个工作表
        for (const sheetName of sheetNames) {
            console.log('\n' + chalk.yellow('━').repeat(50));
            const sheetSpinner = ora({
                text: chalk.blue(`正在处理工作表: ${chalk.bold(sheetName)}`),
                spinner: 'dots'
            }).start();
            
            const worksheet = workbook.Sheets[sheetName];
            
            // 转换为JSON
            const jsonData = xlsx.utils.sheet_to_json(worksheet, {
                defval: null,
                raw: false,
                dateNF: 'YYYY-MM-DD'
            });
            
            // 生成输出路径
            const outputPath = path.join(
                path.dirname(excelPath),
                `${path.basename(excelPath, path.extname(excelPath))}_${sheetName}.json`
            );
            
            // 保存JSON文件
            await fs.promises.writeFile(
                outputPath,
                JSON.stringify(jsonData, null, 2),
                'utf8'
            );
            
            sheetSpinner.succeed(chalk.green(
                `工作表 ${chalk.bold(sheetName)} 转换完成! ` +
                `共 ${chalk.bold(jsonData.length)} 行数据`
            ));
            
            // 移除数据预览部分，直接添加到结果中
            results.push({
                sheetName,
                outputPath,
                rowCount: jsonData.length
            });
        }
        
        return { 
            success: true, 
            results,
            message: `已完成所有工作表的转换，共处理 ${results.length} 个工作表`
        };
        
    } catch (error) {
        return { success: false, error: error.message };
    }
}

// 下载在线文件
async function downloadFile(url) {
    const spinner = ora({
        text: chalk.blue('正在下载在线文件...'),
        spinner: 'dots'
    }).start();

    try {
        const response = await axios({
            method: 'get',
            url: url,
            responseType: 'arraybuffer'
        });

        // 在临时目录创建文件
        const tempDir = path.join(os.tmpdir(), 'excel-to-json');
        if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir);
        }

        const fileName = path.basename(url).split('?')[0] || 'download.xlsx';
        const filePath = path.join(tempDir, fileName);
        
        await fs.promises.writeFile(filePath, response.data);
        spinner.succeed(chalk.green('文件下载完成'));
        
        return filePath;
    } catch (error) {
        spinner.fail(chalk.red('文件下载失败'));
        throw new Error(`下载失败: ${error.message}`);
    }
}

// 检查是否是URL
function isValidUrl(string) {
    try {
        new URL(string);
        return true;
    } catch (_) {
        return false;
    }
}

// 主函数
async function main() {
    console.clear(); // 清屏
    console.log(chalk.bold.cyan('\n📊 Excel转JSON工具'));
    console.log(chalk.dim(`支持的格式: ${chalk.white('.xlsx, .xls, .xlsb, .xlsm, .csv')}`));
    console.log(chalk.yellow('━'.repeat(50)) + '\n');
    
    const promptUser = () => {
        rl.question(chalk.blue('请输入Excel文件路径或URL ') + chalk.dim('(输入q退出): '), async (input) => {
            // 移除输入路径两端的空格和引号
            input = input.trim().replace(/^["']|["']$/g, '');
            
            // 检查是否退出
            if (input.toLowerCase() === 'q') {
                console.log(chalk.green('\n👋 感谢使用，再见!'));
                rl.close();
                return;
            }

            try {
                let excelPath = input;
                
                // 如果是URL，先下载文件
                if (isValidUrl(input)) {
                    excelPath = await downloadFile(input);
                }
                
                // 检查文件是否存在
                if (!fs.existsSync(excelPath)) {
                    console.log(chalk.red(`\n❌ 错误: 找不到文件 '${excelPath}'`));
                    promptUser();
                    return;
                }
                
                // 检查文件扩展名
                const ext = path.extname(excelPath).toLowerCase();
                const supportedExts = ['.xlsx', '.xls', '.xlsb', '.xlsm', '.csv'];
                if (!supportedExts.includes(ext)) {
                    console.log(chalk.red(`\n❌ 错误: 不支持的文件格式。支持的格式: ${supportedExts.join(', ')}`));
                    promptUser();
                    return;
                }
                
                // 执行转换
                const result = await excelToJson(excelPath);
                
                // 如果是下载的临时文件，在转换完成后删除
                if (isValidUrl(input)) {
                    await fs.promises.unlink(excelPath);
                }
                
                if (result.success) {
                    console.log('\n' + chalk.yellow('━'.repeat(50)));
                    console.log(chalk.green(`\n✨ ${result.message}`));
                    console.log(chalk.dim('\n输出文件:'));
                    result.results.forEach(({ sheetName, outputPath, rowCount }) => {
                        console.log(chalk.cyan(`• ${sheetName}: `) + 
                                  chalk.white(`${rowCount} 行数据 → `) + 
                                  chalk.dim(outputPath));
                    });
                } else {
                    console.log(chalk.red(`\n❌ 错误: ${result.error}`));
                }
                
                console.log('\n' + chalk.yellow('━'.repeat(50)) + '\n');
                promptUser();
                
            } catch (error) {
                console.log(chalk.red(`\n❌ 错误: ${error.message}`));
                promptUser();
            }
        });
    };
    
    promptUser();
}

// 处理Ctrl+C
process.on('SIGINT', () => {
    console.log(chalk.yellow('\n\n已取消操作，感谢使用!'));
    rl.close();
});

// 启动程序
main().catch(error => {
    console.error(chalk.red('程序出错:', error));
    rl.close();
});
