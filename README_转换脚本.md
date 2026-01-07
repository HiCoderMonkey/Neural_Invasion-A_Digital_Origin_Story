# Markdown 转 Word 转换脚本使用说明

## 安装依赖

首先需要安装 `python-docx` 库：

```bash
pip install python-docx
```

或者使用 requirements.txt：

```bash
pip install -r requirements.txt
```

## 使用方法∏

### 基本用法

```bash
python md_to_word.py <markdown文件路径>
```

转换后的 Word 文件会自动保存在与 Markdown 文件相同的目录下，文件名相同，扩展名为 `.docx`。

### 指定输出文件

```bash
python md_to_word.py <markdown文件路径> <输出文件路径>
```

## 示例

```bash
# 转换正文文件（输出到同目录下的 正文.docx）
python md_to_word.py 言情版/正文/正文.md

# 指定输出文件名
python md_to_word.py 言情版/正文/正文.md 我的小说.docx
```

## 功能特性

- ✅ 支持 Markdown 标题（#、##、###）
- ✅ 支持粗体文本（**文本**）
- ✅ 支持链接（[文本](URL)），链接文本会保留但不会添加超链接
- ✅ 自动设置中文字体（宋体、黑体）
- ✅ 保持段落格式和空行

## 注意事项

- 脚本会自动处理中文字体设置
- 链接会保留文本内容，但不会在 Word 中添加可点击的超链接
- 如果遇到编码问题，请确保 Markdown 文件使用 UTF-8 编码

