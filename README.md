# Excel 文件对比工具

一个用于对比两个 Excel 文件差异的工具，支持 Web 界面和命令行两种使用方式。

## 功能特性

- 支持多列组合键进行行匹配
- 自动识别仅在单个文件中存在的行
- 逐单元格对比共有行的数据差异
- 生成带颜色标注的差异 Excel 文件
- 支持 Web 界面操作，简单直观
- 支持命令行批量处理
- 可打包为独立可执行文件

## 差异标注说明

- 🟢 绿色背景：仅在当前文件中存在的行
- 🟡 黄色背景 + 红色字体：与另一文件有差异的单元格
- 悬停单元格可查看批注，显示对比文件中的对应值

## 安装依赖

```bash
pip install flask pandas openpyxl python-calamine
```

## 使用方法

### Web 界面

```bash
python app.py
```

启动后访问 `http://127.0.0.1:5000`，按以下步骤操作：

1. 分别上传两个 Excel 文件
2. 选择要对比的 Worksheet
3. 选择用于匹配行的键列（支持多列组合键）
4. 点击"开始对比"
5. 查看对比结果，下载对比报告或标注文件

### 命令行

```bash
python compare_xlsx.py file1.xlsx file2.xlsx --keys 列名1 列名2 --sheet1 Sheet1 --sheet2 Sheet1
```

参数说明：
- `file1`, `file2`：要对比的两个 Excel 文件
- `--keys`：用于匹配行的列名（必填，可指定多个）
- `--sheet1`：文件1的 Worksheet 名称（默认 Sheet1）
- `--sheet2`：文件2的 Worksheet 名称（默认 Sheet1）
- `--output`：输出日志文件路径（可选）

## 打包为可执行文件

```bash
pip install pyinstaller
pyinstaller build.spec
```

打包后的可执行文件位于 `dist` 目录。

## URL 前缀配置

支持通过参数或环境变量配置 URL 前缀，便于反向代理部署：

```bash
python app.py --prefix excel-compare
# 或
set SCRIPT_NAME=excel-compare
python app.py
```

## 项目结构

```
├── app.py              # Web 服务主程序
├── compare_xlsx.py     # 命令行对比工具
├── templates/
│   └── index.html      # Web 界面模板
├── uploads/            # 上传文件临时目录
├── build.spec          # PyInstaller 打包配置
└── README.md
```

## License

MIT
