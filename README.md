# AutoWordTable - Offline Office Word Automatic Table Filling Assistant

A simple and easy-to-use offline tool for automatically filling Word tables based on a custom knowledge base. Previously, I released an open-source application called Autotable (https://github.com/Dukeway/Autotable), which relies on large language models to identify cell values in tables. However, this approach is either inaccurate or requires more powerful models. Autotable uses python-docx, which is cross-platform and simple but does not support merged cells and is suitable only for regular tables. In AutoWordTable, I use the win32com.client method, which can handle merged cells and provides precise control, but it is limited to Windows and requires Microsoft Word to be installed. More importantly, it offers higher recognition accuracy and is completely offline.

## 📝 Features Overview

AutoWordTable is a desktop application that automatically fills tables in Word documents based on data from an Excel knowledge base. The main features include:

- Reading fields and corresponding values from Excel files
- Automatically identifying table fields in Word documents
- Quickly filling tables based on the knowledge base
- Automatically generating and saving the filled documents
- Detailed operation log recording
- Fuzzy matching (optional)
- Filled fields turn red (optional)
- Ignore spaces when matching (optional)
- Multi-row data batch filling: Enable multi-row data batch filling mode (optional, note that the knowledge base table requires multiple fields!)

## 🚀 Usage

1. Run the program and open the main interface.
2. Click the "选择Excel" button to choose the Excel knowledge base containing fields and field values.
3. Click the "选择Word" button to choose the Word template file that needs to be filled.
4. Click the "开始自动表格填充（start）" button to begin processing.
5. The program will generate a "Filled_Table.docx" file in the directory where the Word template is located.

## 🚀 Video Demonstration

video folder

## 📋 Knowledge Base Format Requirements

The Excel knowledge base must include the following columns:
- `Field`: The name of the field to be identified in the table.
- `Field Value`: The value to be filled in for the corresponding field.

Example:

| 字段       | 字段值                 |
|----------|---------------------|
| Name     | Guan Yu             |
| Age      | 42                  |
| Position | Five Tiger Generals |

## 📄 Word Template Format Requirements

The program will automatically identify cells with text in the Word document as field names and fill the corresponding values into the cells to the right of the field names.

## 🔧 Installation Requirements

- Windows operating system
- Python 3.6+
- The following Python libraries:
  - customtkinter
  - pandas
  - pywin32

## 📦 Installation Steps

### Method 1: Using Precompiled Version

1. Download the latest version of the application from my GitHub page.
2. Locate the downloaded file.
3. Run `AutoWordTable.exe`.
4. Select your knowledge base and the table file to be filled.

### Method 2: Installing from Source Code

```bash
# Clone the repository
git clone https://github.com/Dukeway/Autowordtable.git

# Enter the project directory
cd Autowordtable

# Install dependencies
pip install -r requirements.txt

# Run the program
python main.py
```

## 🛠️ Frequently Asked Questions

**Q: Why can't the program recognize my table fields?**  
A: Ensure that the field names in the Excel knowledge base exactly match those in the Word table, including spaces and punctuation.

**Q: What Word document formats are supported?**  
A: Currently, only .docx files are supported.
**Note**: Some tables are in `.doc` files. Even though you have “Save As” the `.doc` file to `.docx`, the content saved in the old `.doc` format (especially tables) may not necessarily be recognized as “standard Word table objects” within Word.

This causes `win32com` to fail to find the “tables” you see when iterating through `.Tables`, and as a result, no operations or logs are generated.
### ✅ **Solution : Force Conversion to Standard Tables**

You can manually standardize the table structure as follows:

1. Open your `.docx` file with Word.

2. Select all the table content and **copy** it.

3. In a new blank `.docx` file, **paste it as “Keep Source Formatting”**.

4. Save the new file.

> This is essentially letting Word “reconstruct” the old-format tables into true `Tables` objects.

**Q: Does it support complex table formats?**  
A: This program is suitable for simple structured tables and partially merged cells. Complex merged cells or nested tables may not be correctly recognized.

**Q: Does it support complex table formats?**  
A: This program is suitable for simple structured tables and partially merged cells. Complex merged cells or nested tables may not be correctly recognized.

## 📜 Open Source License

This project is licensed under the MIT License. For details, see the [LICENSE](LICENSE) file.

## 🙏 Acknowledgements

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - Modern UI component library
- [Pandas](https://pandas.pydata.org/) - Data processing library
- [PyWin32](https://github.com/mhammond/pywin32) - Windows API wrapper

## 📞 Contact Information

If you have any questions or suggestions, please contact the developer via the following methods:
- Email: dukeway@qq.com
- GitHub: [dukeway](https://github.com/your-username)

---

**Note**: This software is open-source and free. Please do not use it for commercial purposes.



# AutoWordTable - 离线的Office Word自动填表助手

一款简单易用的离线的Word表格自动填充工具，可以根据自定义知识库填写docx文件表格。之前我发布了开源的应用：基于大语言模型的自动填表应用Autotable（https://github.com/Dukeway/Autotable）。

但是Autotable依靠模型识别表格中单元格的行列值，要么不准确，要么需要使用到参数更加强大的模型。Autotable使用的是python-docx，支持跨平台、简单，但不支持识别合并单元格，适用于规则表格；我在AutoWordTable中使用win32com.client方法，能处理合并单元格、精确控制，然而仅限 Windows、需要 MS Word 安装。但是更重要的是，识别准确率更高，并且完全离线。


## 📝 功能简介

AutoWordTable是一款桌面应用程序，能够根据Excel知识库中的数据自动填充Word文档中的表格。主要功能包括：

- 从Excel文件读取字段和对应的值
- 自动识别Word文档中的表格字段
- 根据知识库快速填充表格
- 自动生成并保存填充后的文档
- 详细的操作日志记录
- 模糊匹配（可选）
- 填充字段变红（可选）
- 忽略空格匹配（可选）
- 启用多行数据批量填充模式（可选，注意知识库表格需要多个字段！

## 🚀 使用方法

1. 运行程序，打开主界面
2. 点击"选择"按钮，选择包含字段和字段值的Excel知识库
3. 点击"选择"按钮，选择需要填充的Word模板文件
4. 点击"开始自动填表"按钮开始处理
5. 程序将在Word模板所在目录生成"已填写表格.docx"文件

## 🚀 视频演示
在video文件夹中。

## 📋 知识库格式要求

Excel知识库必须包含以下列：
- `字段`: 表格中需要识别的字段名称
- `字段值`: 对应字段的填充值

示例：

| 字段 | 字段值  |
|------|------|
| 姓名 | 关羽   |
| 年龄 | 42   |
| 职位 | 五虎上将 |

## 📄 Word模板格式要求

程序会自动识别Word文档有文字存在的单元格内容作为字段名，并将对应的值填入字段名称右侧的单元格中。

## 🔧 安装要求

- Windows操作系统
- Python 3.6+ 
- 以下Python库:
  - customtkinter
  - pandas
  - pywin32

## 📦 安装步骤

### 方法1: 使用预编译版本

1. 从我的Github页面下载Release最新版本应用程序
2. 找到文件下载的位置
3. 运行 `AutoWordTable.exe`
4. 选择你的知识库和需要填写的表格文件。

### 方法2: 从源码安装

```bash
# 克隆仓库
git clone https://github.com/Dukeway/Autowordtable.git

# 进入项目目录
cd Autowordtable

# 安装依赖
pip install -r requirements.txt

# 运行程序
python main.py
```

## 🛠️ 常见问题

**Q: 为什么程序无法识别我的表格字段?**  
A: 请确保Excel知识库中的字段名称与Word表格中的完全一致，包括空格和标点符号。

**Q: 程序支持哪些Word文档格式?**  
A: 目前仅支持.docx格式的文件。
注意，某些表格是doc文件，虽然你把 `.doc` 文件“另存为”成了 `.docx`，但 **老版本 `.doc` 格式保存的内容（尤其是表格）在 Word 内部并不一定被识别为“标准的 Word 表格对象”**。

这就导致 `win32com` 在遍历 `.Tables` 时，**没有找到你看到的“表格”**，也就没有任何操作或日志输出。

### ✅ 解决方法：强制转换为标准表格

你可以用如下步骤**手动标准化表格结构**：

1. 用 Word 打开你的 `.docx` 文件

2. 选中所有表格内容，**复制**

3. 在一个新建的空白 `.docx` 文件中，**粘贴为“保留源格式”**

4. 保存新文件

> 这相当于让 Word 把老旧格式的表格“重构”为真正的 `Tables` 对象。

**Q: 是否支持复杂的表格格式?**  
A: 本程序适用于结构简单的表格以及部分合并的单元格，复杂的合并单元格或嵌套表格可能无法正确识别。

## 📜 开源许可

本项目采用MIT许可证，详情请查看[LICENSE](LICENSE)文件。

## 🙏 鸣谢

- [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter) - 现代化UI组件库
- [Pandas](https://pandas.pydata.org/) - 数据处理库
- [PyWin32](https://github.com/mhammond/pywin32) - Windows API封装

## 📞 联系方式

如有问题或建议，请通过以下方式联系开发者：
- Email: dukeway@qq.com
- GitHub: [dukeway](https://github.com/your-username)

---

**注意**: 本软件为开源免费软件，请勿用于商业用途。
