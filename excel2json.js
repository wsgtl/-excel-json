const XLSX = require('xlsx');
const fs = require('fs-extra');
const path = require('path');
const yargs = require('yargs');
const chalk = require('chalk');

class ExcelToJsonConverter {
    constructor() {
        this.supportedFormats = ['.xlsx', '.xls'];
    }

    /**
     * 转换指定目录中的所有 Excel 文件
     * @param {string} inputDir 输入目录（包含 Excel 文件）
     * @param {string} outputDir 输出目录
     * @param {Object} options 配置选项
     */
    async convertDirectory(inputDir, outputDir, options = {}) {
        try {
            // 检查输入目录是否存在
            if (!fs.existsSync(inputDir)) {
                throw new Error(`输入目录不存在: ${inputDir}`);
            }

            // 创建输出目录
            await fs.ensureDir(outputDir);

            // 查找所有 Excel 文件
            const files = fs.readdirSync(inputDir);
            const excelFiles = files.filter(file => 
                this.supportedFormats.includes(path.extname(file).toLowerCase())
            );

            if (excelFiles.length === 0) {
                console.log(chalk.yellow('⚠️  未找到 Excel 文件'));
                return { success: 0, total: 0 };
            }

            console.log(chalk.blue(`📁 找到 ${excelFiles.length} 个 Excel 文件`));

            let successCount = 0;
            const results = {};

            // 处理每个 Excel 文件
            for (const excelFile of excelFiles) {
                const excelPath = path.join(inputDir, excelFile);
                const fileName = path.basename(excelFile, path.extname(excelFile));
                
                console.log(chalk.cyan(`\n🔄 处理文件: ${excelFile}`));

                try {
                    const workbook = XLSX.readFile(excelPath);
                    const sheetNames = workbook.SheetNames;

                    // 处理每个工作表
                    for (const sheetName of sheetNames) {
                        const worksheet = workbook.Sheets[sheetName];
                        
                        // 获取原始数据
                        const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                            header: 1, // 使用数组格式
                            defval: '',
                            raw: false  // 允许类型转换
                        });

                        if (jsonData.length === 0) {
                            console.log(chalk.yellow(`  ⚠️  工作表 ${sheetName} 为空，跳过`));
                            continue;
                        }

                        console.log(chalk.gray(`  📊 原始数据: ${JSON.stringify(jsonData)}`));

                        // 检测数据结构并转换
                        const convertedData = this.detectAndConvertStructure(jsonData, sheetName);
                        
                        // 生成输出文件名
                        const outputFileName = sheetNames.length > 1 ? 
                            `${fileName}_${this.sanitizeFileName(sheetName)}.json` : 
                            `${fileName}.json`;
                        
                        const outputPath = path.join(outputDir, outputFileName);
                        
                        // 保存 JSON 文件
                        await fs.writeJson(outputPath, convertedData, { spaces: 2 });
                        
                        const recordCount = Array.isArray(convertedData) ? convertedData.length : Object.keys(convertedData).length;
                        const structureType = Array.isArray(convertedData) ? '数组' : '键值对';
                        console.log(chalk.green(`  ✅ 生成: ${outputFileName} (${structureType}, ${recordCount} 条记录)`));
                        
                        results[outputFileName] = {
                            recordCount,
                            structureType,
                            source: `${excelFile}/${sheetName}`
                        };
                    }

                    successCount++;

                } catch (error) {
                    console.log(chalk.red(`  ❌ 处理文件 ${excelFile} 失败: ${error.message}`));
                }
            }

            console.log(chalk.green(`\n🎉 转换完成！成功: ${successCount}/${excelFiles.length} 个文件`));
            return { 
                success: successCount, 
                total: excelFiles.length,
                results 
            };

        } catch (error) {
            console.log(chalk.red(`❌ 转换过程出错: ${error.message}`));
            throw error;
        }
    }

    /**
     * 检测数据结构并转换
     */
    detectAndConvertStructure(data, sheetName) {
        if (data.length === 0) return {};
        
        const firstRowFirstCell = data[0] && data[0][0];
        
        console.log(chalk.gray(`  🔍 检测数据结构，第一行第一列: "${firstRowFirstCell}"`));

        // 根据第一行第一列的值判断结构类型
        if (this.isKeyValueStructure(firstRowFirstCell)) {
            console.log(chalk.blue('  🔑 检测为键值对结构'));
            return this.convertKeyValueStructure(data);
        } else {
            console.log(chalk.blue('  📋 检测为数组结构'));
            return this.convertArrayStructure(data);
        }
    }

    /**
     * 检测是否为键值对结构
     */
    isKeyValueStructure(firstCell) {
        if (!firstCell) return false;
        
        const firstCellStr = String(firstCell).toLowerCase().trim();
        
        // 如果第一行第一列是 "key" 或包含 "key" 关键字，则是键值对结构
        return firstCellStr === 'key' || firstCellStr.includes('key');
    }

    /**
     * 转换键值对结构
     */
    convertKeyValueStructure(data) {
        const result = {};
        
        console.log(chalk.gray(`  🔄 开始转换键值对结构，共 ${data.length} 行`));

        // 跳过第一行（标题行），从第二行开始
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            
            // 跳过空行
            if (!row || row.length < 2 || row.every(cell => cell === '' || cell === null || cell === undefined)) {
                continue;
            }

            // 第一列是key，第二列是value
            const key = row[0];
            const value = row[1];
            
            if (key !== undefined && key !== '' && key !== null) {
                const processedKey = this.processKey(key);
                const processedValue = this.processValue(value, key);
                
                result[processedKey] = processedValue;
                console.log(chalk.gray(`    ${processedKey} = ${JSON.stringify(processedValue)}`));
            }
        }
        
        console.log(chalk.gray(`  ✅ 键值对转换完成，共 ${Object.keys(result).length} 个键值对`));
        return result;
    }

    /**
     * 转换数组结构
     */
    convertArrayStructure(data) {
        const result = [];
        
        if (data.length < 2) return result;
        
        const headers = data[0];
        
        console.log(chalk.gray(`  📝 标题行: ${JSON.stringify(headers)}`));

        // 从第二行开始处理数据（跳过标题行）
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const item = {};
            
            // 跳过空行
            if (!row || row.every(cell => cell === '' || cell === null || cell === undefined)) {
                continue;
            }
            
            for (let j = 0; j < headers.length; j++) {
                const key = headers[j];
                const value = row[j];
                
                if (key === undefined || key === '' || key === null) continue;
                
                const processedKey = this.processKey(key);
                item[processedKey] = this.processValue(value, key);
            }
            
            // 只有有数据的行才加入结果
            if (Object.keys(item).length > 0) {
                result.push(item);
            }
        }
        
        console.log(chalk.gray(`  ✅ 数组转换完成，共 ${result.length} 条记录`));
        return result;
    }

    /**
     * 处理键名
     */
    processKey(key) {
        if (typeof key !== 'string') return String(key);
        return key.trim();
    }

    /**
     * 处理值
     */
    processValue(value, key) {
        if (value === '' || value === null || value === undefined) {
            return null;
        }
        
        // 如果是字符串，进行修剪
        if (typeof value === 'string') {
            value = value.trim();
            if (value === '') return null;
        }
        
        // 尝试解析 JSON 字符串
        if (typeof value === 'string') {
            // 处理数组字符串
            if (value.startsWith('[') && value.endsWith(']')) {
                try {
                    return JSON.parse(value);
                } catch (e) {
                    // 如果不是合法 JSON，按逗号分割
                    if (value.includes(',')) {
                        const items = value.slice(1, -1).split(',').map(item => item.trim()).filter(item => item);
                        return items.length > 0 ? items : null;
                    }
                }
            }
            
            // 处理对象字符串
            if (value.startsWith('{') && value.endsWith('}')) {
                try {
                    return JSON.parse(value);
                } catch (e) {
                    // 解析失败，返回原字符串
                }
            }
            
            // 尝试转换为数字
            if (!isNaN(value) && value !== '') {
                const num = Number(value);
                if (!isNaN(num)) return num;
            }
            
            // 处理布尔值
            const lowerValue = value.toLowerCase();
            if (lowerValue === 'true' || lowerValue === 'false') {
                return lowerValue === 'true';
            }
        }
        
        return value;
    }

    /**
     * 清理文件名
     */
    sanitizeFileName(name) {
        return name.replace(/[\\/*?:"<>|]/g, '_');
    }
}

// CLI 接口
if (require.main === module) {
    const argv = yargs
        .usage('用法: $0 <command> [选项]')
        .command('convert', '转换 Excel 文件为 JSON', {
            input: {
                alias: 'i',
                describe: 'Excel 文件所在目录',
                demandOption: true,
                type: 'string'
            },
            output: {
                alias: 'o',
                describe: 'JSON 输出目录',
                demandOption: true,
                type: 'string'
            },
            raw: {
                describe: '保留原始值',
                type: 'boolean',
                default: false
            }
        })
        .example('$0 convert -i ./excels -o ./jsons', '转换 excels 目录中的所有 Excel 文件')
        .help('h')
        .alias('h', 'help')
        .argv;

    if (argv._[0] === 'convert') {
        const converter = new ExcelToJsonConverter();
        converter.convertDirectory(argv.input, argv.output, { raw: argv.raw })
            .then(() => process.exit(0))
            .catch(() => process.exit(1));
    } else {
        yargs.showHelp();
    }
}

module.exports = ExcelToJsonConverter;