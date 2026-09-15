# 墨转 · Markdown 转 Word

一个面向中文写作的 Streamlit 小网站：在浏览器里粘贴 Markdown、实时预览，并导出 `.docx`。

## 关键能力

- 普通段落使用 Word 中名为“正文”的样式，默认宋体、五号（10.5 pt），A4 页面；
- `$...$`、`$$...$$`、`\\(...\\)`、`\\[...\\]` 公式导出为 Word 原生 OMML，可继续编辑；
- 自动识别 `latex` 代码块和复制文本中常见的双重反斜杠，下载后直接显示为 Word 专业格式；
- 支持标题、列表、引用、代码块、脚注和 Markdown 表格；
- 每次导出后自动检查 DOCX 完整性、字体字号和原生公式节点；
- 文稿只在临时目录中完成转换，不写入数据库。

## 本地运行

建议使用 Python 3.11 或 3.12。

```bash
python -m venv .venv
# Windows
.venv\\Scripts\\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

`pypandoc-binary` 已自带 Pandoc，因此通常无需额外安装系统软件。

## 测试

```bash
pip install -r dev-requirements.txt
pytest -q
```

测试会实际生成 Word，并检查 `word/document.xml` 中的结构化 OMML 节点、线性 LaTeX 残留、宋体样式和五号字号。

## 发布到 GitHub

在本目录运行：

```bash
git init
git add .
git commit -m "feat: build Markdown to editable Word converter"
git branch -M main
git remote add origin https://github.com/liang3445132/md2word-streamlit.git
git push -u origin main
```

## 部署到 Streamlit Community Cloud

1. 打开 [share.streamlit.io](https://share.streamlit.io/) 并使用 GitHub 登录；
2. 选择 **Create app**，选中刚才的仓库与 `main` 分支；
3. Main file path 填 `app.py`；
4. 点击 Deploy。

建议仓库名使用 `md2word-streamlit`，Main file path 直接填写 `app.py`。

本项目不需要 Secrets。首次构建需要下载安装依赖，通常会比后续重启稍慢。

## 技术说明

转换链路为：

```text
Markdown + LaTeX
        ↓ Pandoc
DOCX + OMML 原生公式
        ↓ python-docx
宋体五号正文 + A4 页面 + OOXML 自检
```

公式不是截图或 SVG，因此能在 Microsoft Word 的公式编辑器中继续修改。少数依赖自定义宏或 LaTeX 宏包的命令，需要先改写为 Pandoc 支持的标准 LaTeX。
