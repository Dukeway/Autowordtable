# AutoWordTable - Office Word自动填表助手

一款简单易用的Word表格自动填充工具，由Dukeway Zhong个人开发的开源免费软件。



## 📝 功能简介

AutoWordTable是一款桌面应用程序，能够根据Excel知识库中的数据自动填充Word文档中的表格。主要功能包括：

- 从Excel文件读取字段和对应的值
- 自动识别Word文档中的表格字段
- 根据知识库快速填充表格
- 自动生成并保存填充后的文档
- 详细的操作日志记录

## 🚀 使用方法

1. 运行程序，打开主界面
2. 点击"选择"按钮，选择包含字段和字段值的Excel知识库
3. 点击"选择"按钮，选择需要填充的Word模板文件
4. 点击"开始自动填表"按钮开始处理
5. 程序将在Word模板所在目录生成"已填写表格.docx"文件

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

程序会自动识别Word文档中表格的第一列作为字段名，并将对应的值填入字段名称右侧的单元格中。

## 🔧 安装要求

- Windows操作系统
- Python 3.6+ 
- 以下Python库:
  - customtkinter
  - pandas
  - pywin32

## 📦 安装步骤

### 方法1: 使用预编译版本

1. 从[发布页面](https://github.com/your-username/autowordtable/releases)下载最新版本
2. 解压缩文件
3. 运行 `AutoWordTable.exe`

### 方法2: 从源码安装

```bash
# 克隆仓库
git clone https://github.com/Dukeway/autowordtable.git

# 进入项目目录
cd autowordtable

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

**Q: 是否支持复杂的表格格式?**  
A: 本程序适用于结构简单的表格，复杂的合并单元格或嵌套表格可能无法正确识别。

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