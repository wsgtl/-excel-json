const fs = require('fs-extra');
const path = require('path');
const chalk = require('chalk');

class ProjectGenerator {
    constructor() {
        this.projectsRoot = path.join(__dirname, 'projects');
    }

    /**
     * 创建新项目
     * @param {string} projectName 项目名称
     */
    async createProject(projectName) {
        try {
            const projectPath = path.join(this.projectsRoot, projectName);
            const excelsPath = path.join(projectPath, 'excels');
            const jsonsPath = path.join(projectPath, 'jsons');

            // 创建目录结构
            await fs.ensureDir(projectPath);
            await fs.ensureDir(excelsPath);
            await fs.ensureDir(jsonsPath);

            // 创建 convert.bat 文件
            const batContent = this.generateBatFile(projectName);
            const batPath = path.join(projectPath, 'convert.bat');
            await fs.writeFile(batPath, batContent);

            // 创建示例 Excel 文件
            await this.createExampleFiles(excelsPath);

            console.log(chalk.green(`🎉 项目 "${projectName}" 创建成功！`));
            console.log(chalk.blue(`📁 项目路径: ${projectPath}`));
            console.log(chalk.cyan(`📝 使用方法:`));
            console.log(`  1. 将 Excel 文件放入 ${chalk.yellow('excels/')} 目录`);
            console.log(`  2. 双击运行 ${chalk.yellow('convert.bat')}`);
            console.log(`  3. 查看生成的 JSON 文件在 ${chalk.yellow('jsons/')} 目录`);

            return projectPath;

        } catch (error) {
            console.log(chalk.red(`❌ 创建项目失败: ${error.message}`));
            throw error;
        }
    }

    /**
     * 生成批处理文件内容
     */
    generateBatFile(projectName) {
        return `@echo off
chcp 65001 >nul
echo ===============================================
echo  Excel 转 JSON 转换工具 - ${projectName}
echo ===============================================
echo.

cd /d "%~dp0"

if not exist "excels" (
    echo ❌ 错误: 未找到 excels 目录
    pause
    exit /b 1
)

echo 🔄 开始转换 Excel 文件...
node "..\\..\\excel2json.js" convert -i "excels" -o "jsons"

if %errorlevel% equ 0 (
    echo.
    echo ✅ 转换完成！
    echo 📁 JSON 文件已保存到 jsons 目录
) else (
    echo.
    echo ❌ 转换失败！
)

echo.
pause
`;
    }

    /**
     * 创建示例文件
     */
    async createExampleFiles(excelsPath) {
        const exampleContent = `
示例 Excel 文件结构说明:

1. 普通 Key-Value 结构 (config.xlsx):
   | key       | value     |
   |-----------|-----------|
   | game_name | 我的游戏  |
   | version   | 1.0.0     |

2. 数组结构 (items.xlsx):
   | id | name  | type    | value |
   |----|-------|---------|-------|
   | 1  | 金币  | currency| 100   |
   | 2  | 钻石  | currency| 50    |

3. 包含数组字段 (levels.xlsx):
   | level | rewards[]    | multiplier |
   |-------|--------------|------------|
   | 1     | [coin,gem]   | 1.5        |
   | 2     | [gem,chest]  | 2.0        |

将您的 Excel 文件放入此目录，然后运行 convert.bat 进行转换。
        `.trim();

        await fs.writeFile(path.join(excelsPath, 'README.txt'), exampleContent);
    }

    /**
     * 列出所有项目
     */
    async listProjects() {
        if (!fs.existsSync(this.projectsRoot)) {
            return [];
        }

        const items = await fs.readdir(this.projectsRoot);
        const projects = [];

        for (const item of items) {
            const itemPath = path.join(this.projectsRoot, item);
            const stat = await fs.stat(itemPath);
            
            if (stat.isDirectory()) {
                projects.push(item);
            }
        }

        return projects;
    }
}

// CLI 接口
if (require.main === module) {
    const yargs = require('yargs');

    const argv = yargs
        .usage('用法: $0 <command> [选项]')
        .command('new <name>', '创建新项目', {
            name: {
                describe: '项目名称',
                demandOption: true,
                type: 'string'
            }
        })
        .command('list', '列出所有项目')
        .example('$0 new my-game', '创建名为 my-game 的新项目')
        .example('$0 list', '列出所有现有项目')
        .help('h')
        .alias('h', 'help')
        .argv;

    const generator = new ProjectGenerator();

    if (argv._[0] === 'new') {
        generator.createProject(argv.name);
    } else if (argv._[0] === 'list') {
        generator.listProjects().then(projects => {
            if (projects.length === 0) {
                console.log(chalk.yellow('暂无项目'));
            } else {
                console.log(chalk.blue('现有项目:'));
                projects.forEach(project => {
                    console.log(`  📁 ${project}`);
                });
            }
        });
    } else {
        yargs.showHelp();
    }
}

module.exports = ProjectGenerator;