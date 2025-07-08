# Excel数据清洗合并工具

## 项目简介
本工具是一款基于Electron的极简桌面应用，专为批量处理结构复杂的Excel文件而设计。支持多文件、多Sheet自动识别与清洗，界面全中文，操作简单，适合在Windows环境下便携运行（绿色版/免安装）。

## 主要功能
- **批量处理**：支持选择文件夹，一键清洗合并所有Excel（.xlsx）文件。
- **智能识别**：自动识别多表头、跨行表头、sheet内多个“界址点成果表”。
- **字段提取**：仅保留“点号、x(m)、y(m)”三列，兼容不同表头写法。
- **去重逻辑**：sheet内xy坐标相同只保留一条（J1点闭合时首尾J1均保留），多文件合并时不去重。
- **日志详细**：实时显示每个文件、每个表的处理进度与去重详情。
- **极简界面**：全中文，输入/输出文件夹选择、日志窗口、操作按钮一目了然。
- **绿色便携**：打包后为单文件exe，无需安装，适合内网分发。

## 使用方法
### 1. 环境准备
- Windows 10/11
- [Node.js](https://nodejs.org/) 16+（仅打包或开发时需要，普通用户直接用exe即可）

### 2. 界面操作
1. 双击`release/excel-cleaner-win32-x64/excel-cleaner.exe`启动程序。
2. 点击“选择文件夹”按钮，分别选择**输入文件夹**（含待处理Excel）和**输出文件夹**（保存结果）。
3. 点击“开始处理”，日志窗口会实时显示详细进度。
4. 处理完成后，输出文件夹下会生成`清洗合并结果.xlsx`。
5. 可随时点击“清空日志”清除窗口内容。

### 3. 批量处理示例
- 输入文件夹示例：
  - `00/局部一.xlsx`
  - `00/区块二.xlsx`
  - `00/区块三.xlsx`
- 输出文件示例：
  - `00/清洗合并结果.xlsx`

### 4. 开发与打包
如需自行打包或二次开发：
```powershell
# 安装依赖
npm install

# 开发调试
npm run start

# 构建前端（如需React版本）
npm run build

# 打包为绿色版exe（推荐使用国内镜像加速）
$env:ELECTRON_MIRROR="https://npmmirror.com/mirrors/electron/"
npx electron-packager . excel-cleaner --platform=win32 --arch=x64 --out=release --overwrite --electron-version=37.2.0 --main=main.js
```

## 依赖环境
- electron ^37.2.0
- xlsx ^0.18.5
- @electron/remote ^2.1.3
- react ^19.1.0（如用React界面）

## 常见问题与注意事项
- **打包慢/下载失败**：建议设置国内镜像（如npmmirror），或手动下载Electron二进制。
- **icon.ico缺失警告**：不影响程序运行，可忽略。
- **双package.json问题**：本项目已规范为主目录单一配置。
- **数据量大**：已支持单表10000行，极端大文件建议分批处理。
- **仅支持.xlsx格式**，如有.xls请先转换。

## 目录结构示例
```
├─00
│   ├─局部一.xlsx
│   ├─区块二.xlsx
│   ├─区块三.xlsx
│   └─清洗合并结果.xlsx
├─main.js
├─renderer-native.js
├─index-native.html
├─package.json
└─...
```

## 联系与反馈
如有问题或建议，欢迎通过Issue或邮件联系作者。 