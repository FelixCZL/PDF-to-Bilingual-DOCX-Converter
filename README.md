# PDF to Bilingual DOCX Converter

一个强大的 Python 工具，可以将 PDF 文件转换为中英双语 Word 文档。支持文本提取和 OCR 识别，自动翻译英文内容为中文。

## 功能特性

- 📄 **PDF 文本提取**：自动提取 PDF 中的文本内容
- 🔍 **OCR 识别**：支持扫描版 PDF，使用 Tesseract OCR 进行文字识别
- 🌐 **自动翻译**：使用 Google 翻译 API 将英文内容翻译为中文
- 📝 **双语文档**：生成格式化的中英双语 Word 文档
- 🔄 **批量处理**：自动处理当前目录下的所有 PDF 文件

## 系统要求

- Python 3.7+
- Tesseract OCR（用于 OCR 功能）

### 安装 Tesseract OCR

**Windows:**
1. 下载安装包：https://github.com/UB-Mannheim/tesseract/wiki
2. 安装到默认路径：`C:\Program Files\Tesseract-OCR\`
3. 或安装到其他路径后，修改代码中的路径配置

**macOS:**
```bash
brew install tesseract
brew install tesseract-lang  # 安装中文语言包
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
sudo apt-get install tesseract-ocr-chi-sim  # 安装中文语言包
```

## 安装

1. 克隆或下载此仓库

2. 安装 Python 依赖：
```bash
pip install -r requirements.txt
```

## 使用方法

1. 将需要转换的 PDF 文件放在脚本所在目录

2. 运行脚本：
```bash
python pdf_to_bilingual_docx.py
```

3. 脚本会自动：
   - 扫描当前目录下的所有 PDF 文件
   - 提取文本内容（支持 OCR）
   - 翻译英文内容为中文
   - 生成对应的 `.docx` 文件

## 输出格式

生成的 Word 文档包含：
- 文档标题（基于 PDF 文件名）
- 每页内容，包含：
  - 英文原文
  - 中文翻译
  - 段落分隔符

对于通过 OCR 识别的页面，会特别标注。

## 配置

### Tesseract OCR 路径

如果 Tesseract 安装在非默认路径，可以修改代码中的 `setup_tesseract()` 函数。

### 翻译设置

默认使用 Google 翻译，目标语言为简体中文。可以在 `translate_text()` 函数中修改：
```python
translator = GoogleTranslator(source='auto', target='zh-CN')
```

## 注意事项

- 翻译功能依赖网络连接
- 大量文本翻译可能需要较长时间
- OCR 识别准确度取决于 PDF 图片质量
- 建议在使用前备份原始 PDF 文件

## 依赖库

- `PyMuPDF` - PDF 处理
- `python-docx` - Word 文档生成
- `deep-translator` - 文本翻译
- `Pillow` - 图像处理
- `pytesseract` - OCR 识别

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request！

## 更新日志

### v1.0.0
- 初始版本
- 支持 PDF 文本提取
- 支持 OCR 识别
- 支持自动翻译
- 支持批量处理

