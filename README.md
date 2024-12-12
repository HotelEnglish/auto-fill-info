# Doc Auto Fill

一个自动填写 Word 文档表格的工具。

## 功能特点

- 支持批量处理多个文档（最多10个）
- 支持 .doc 和 .docx 格式
- 智能字段匹配
- 图形用户界面
- 保持原文档格式

## 安装要求

- Windows 操作系统
- Microsoft Word
- Python 3.8 或更高版本

## 安装步骤

### 1. 克隆仓库：
```bash
git clone https://github.com/your-username/doc-auto-fill.git
```

### 2. 安装依赖：

```bash
pip install -r requirements.txt
```

## 使用方法

1. 准备个人信息文件（.docx格式），格式如下：
```
姓名: 张三
性别: 男
出生日期: 1990年1月1日
...
```

2. 运行程序：

```bash
python src/auto_fill_gui.py
```

3. 按照界面提示操作：
   - 选择个人信息文件
   - 添加需要填写的文档（最多10个）
   - 点击"开始填写"

## 示例

查看 `examples` 目录获取示例文件。

## 贡献指南

欢迎提交 Pull Request 或创建 Issue。

## 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。















