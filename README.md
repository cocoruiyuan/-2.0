# Russian Handwriting Generator

一个使用 Streamlit 制作的俄语手写图片生成器。

## 功能

- 粘贴俄语文字
- 上传 TXT
- 上传 DOCX
- 上传文字型 PDF
- 自动换行
- 自动分页
- 横线纸、白纸、方格纸
- 蓝色或黑色墨水
- 手写随机扰动
- 导出第一页 PNG
- 导出全部 PNG ZIP
- 导出多页 PDF

## 项目结构

```text
russian-handwriting-app/
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── .streamlit/
│   └── config.toml
└── fonts/
    └── your-russian-handwriting-font.ttf
```

## 字体

你必须自己放入一个支持俄语西里尔字母的 `.ttf` 或 `.otf` 手写字体。

把字体放到：

```text
fonts/
```

建议先测试下面这句话：

```text
Съешь ещё этих мягких французских булок, да выпей чаю.
```

注意：请确认你有权在 GitHub 项目或公开网站中使用和分发该字体。

## 本地运行

### 1. 创建虚拟环境

Windows：

```bash
python -m venv .venv
.venv\Scripts\activate
```

macOS / Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 2. 安装依赖

```bash
python -m pip install -r requirements.txt
```

### 3. 启动

```bash
python -m streamlit run app.py
```

浏览器打开：

```text
http://localhost:8501
```

## 当前限制

- 扫描版 PDF 暂不支持 OCR。
- 旧版 `.doc` 暂不支持，请先另存为 `.docx`。
- 复杂表格和原始 Word 排版不会完全保留。
- 当前主要通过手写字体和随机扰动模拟真实笔迹。
