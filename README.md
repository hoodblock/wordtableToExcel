# Word表格转换成Excel列表，可配置自己想要的列表数据

---
## 📂 使用说明
目前开发适用于特定word文档转特定excel，是适用于特定工作要求而完成的，基础功能都已实现，可对项目进行二次开发，完成通用适配
把项目中`wordTable_0-4.docx`文件拖入项目界面可以查看项目的整个工作流程

## 📂 界面样例


<img width="495" alt="截屏2025-06-12 17 49 44" src="https://github.com/user-attachments/assets/fc39b196-dbc2-4711-aaaa-52a398e6e336" />


<img width="494" alt="截屏2025-06-12 17 51 18" src="https://github.com/user-attachments/assets/5f48af76-300f-4b6e-9f92-244569398e33" />

---
## ✨ 功能特点
- ✅ 从电脑选择Word文件或者拖拽文件到软件工程
- ✅ 自动提取 Word 文档中的所有表格生成二维数组
- ✅ 支持一个或多个表格生成多个 Excel Sheet
- ✅ 列表宽度调整、支持行样式（如分隔行背景色）
- ✅ 支持选择导出路径，导出 `.xlsx` 文件

## 📂 项目结构说明

📁 WordToExcelApp

├── wordTable_0.docx             // 特定的word表格，用此来生成excel表格数据

├── ContentView.swift            // 软件 主界面

├── WordTableParser.swift        // Word 表格解析逻辑（XML 解压）

├── WordTestCaseParser.swift     // 封装word数据并生成特定的表格

├── ExcelHeader.swift            // 自定义表头结构配置

├── SavePanelHelper.swift        // NSSavePanel 路径选择封装

## 🧩 技术说明

#### 📄 Word 表格解析
.docx 实质是一个 ZIP 压缩包，表格结构在 word/document.xml 中，可使用 Swift 原生 ` XMLParser ` 解析 ` <w:tbl>，<w:tr>，<w:tc>，<w:t> `  表格节点

#### 📊 Excel 表格生成
使用 ` libxlsxwriter ` Swift 包，支持单元格样式、列宽、字体加粗、背景色



