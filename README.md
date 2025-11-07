# Excel 转 JSON 工具

一个简单易用的 Excel 转 JSON 转换工具，专为游戏配置数据和应用程序静态数据设计。

## 🚀 快速开始

### 环境要求
- **Node.js** (版本 ≥ 14.0.0)
  - 下载地址: https://nodejs.org/
  - 安装时勾选 "Add to PATH"

### 安装步骤
1. 下载工具文件到 `excel-to-json-tool` 目录
2. 安装依赖:
   ```bash
   npm install
###  使用方法
1. 创建项目: 运行 create-project.bat，输入项目名称
1. 准备数据: 将 Excel 文件放入 projects/项目名/excels/ 目录
1. 一键转换: 双击 projects/项目名/convert.bat
1. 获取结果: 在 projects/项目名/jsons/ 目录查看生成的 JSON 文件

### Excel 表格格式
#### 键值对结构 (第一行第一列为 "key")
    ```text
    key         value
    game_name   黄金聚宝盆
    version     1.0.0
    max_players 1000
    is_active   true
    gg[]        2     3     6     8 
    ```
##### 输出 JSON:

    ```json
    {"game_name":"黄金聚宝盆","version":"1.0.0","max_players":1000,"is_active":true,"gg":[2,3,6,8]}
    ```
##### 说明:

第一行标题会被跳过

键名以 [] 结尾的字段会转换为数组

支持自动类型转换 (数字、布尔值)

#### 数组结构
第一行第一列是 "id" 或其他值，自动跳过注释行

    ```text
    id  name    num
    1   hhh     22
    2   fdf     23433
    3   hhh     24
    ```

#### 输出：

    ```json
    [{"id":1,"name":"hhh","num":22},{"id":2,"name":"fdf","num":23433},{"id":3,"name":"hhh","num":24}]
    ```

### 目录结构
    ```text
    excel-to-json-tool/
    ├── package.json
    ├── excel2json.js
    ├── project-generator.js
    ├── create-project.bat
    └── projects/
        └── 项目名称/
            ├── convert.bat
            ├── excels/
            └── jsons/
    ```

### 批处理文件说明
#### create-project.bat
* 位置：工具根目录

* 功能：创建新项目目录

* 用法：双击运行，输入项目名称

#### convert.bat
* 位置：各项目目录内

* 功能：转换该项目下的 Excel 文件

* 用法：双击运行，自动生成 JSON 文件
