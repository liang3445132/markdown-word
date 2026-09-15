from __future__ import annotations

import hashlib
import re

import streamlit as st

from converter import ConversionError, convert_markdown_to_docx, inspect_docx


st.set_page_config(
    page_title="墨转 · Markdown 转 Word",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="collapsed",
)


EXAMPLE_MARKDOWN = r"""# Markdown 转 Word 示例

这是一段普通正文。导出后默认使用 **宋体、五号（10.5 pt）**，适合课程作业、实验报告和论文初稿。

## 可编辑公式

行内公式：质能方程 $E=mc^2$，以及欧拉恒等式 $e^{i\pi}+1=0$。

独立公式：

$$
\int_{-\infty}^{+\infty} e^{-x^2}\,dx=\sqrt{\pi}
$$

矩阵与分式：

$$
A=\begin{bmatrix}1&2\\3&4\end{bmatrix},\qquad
\hat{\beta}=(X^TX)^{-1}X^Ty
$$

## 表格、列表与引用

| 功能 | 导出结果 |
|---|---|
| 正文 | 宋体五号 |
| 公式 | Word 原生公式，可双击编辑 |
| 表格 | Word 表格 |

1. 在左侧粘贴 Markdown；
2. 在右侧检查预览；
3. 点击生成并下载 Word。

> 提示：公式请使用 `$...$` 或 `$$...$$` 包裹。
"""


def load_example() -> None:
    st.session_state.md_content = EXAMPLE_MARKDOWN


def clear_editor() -> None:
    st.session_state.md_content = ""


def safe_markdown_name(name: str) -> str:
    stem = re.sub(r"\.(md|markdown|txt)$", "", name, flags=re.IGNORECASE)
    stem = re.sub(r"[^\w\u4e00-\u9fff-]+", "-", stem).strip("-_")
    return stem[:60] or "markdown-document"


if "md_content" not in st.session_state:
    st.session_state.md_content = EXAMPLE_MARKDOWN


st.markdown(
    """
<style>
    :root {
        --ink: #18211c;
        --muted: #647069;
        --paper: #fffefa;
        --leaf: #2f6b4f;
        --leaf-soft: #e8f1eb;
        --line: #dfe6e1;
    }
    [data-testid="stAppViewContainer"] {
        background:
          radial-gradient(circle at 12% 5%, rgba(221, 236, 225, .72), transparent 28rem),
          linear-gradient(180deg, #f7faf7 0%, #f3f6f3 100%);
        color: var(--ink);
    }
    [data-testid="stHeader"] { background: transparent; }
    [data-testid="stMainBlockContainer"] {
        max-width: 1440px;
        padding-top: 1.3rem;
        padding-bottom: 3rem;
    }
    #MainMenu, footer { visibility: hidden; }
    .brand {
        display: flex; align-items: center; gap: .65rem;
        color: var(--leaf); font-size: .9rem; font-weight: 750;
        letter-spacing: .08em; margin-bottom: 2.5rem;
    }
    .brand-mark {
        display: grid; place-items: center; width: 2rem; height: 2rem;
        border-radius: .65rem; color: white; background: var(--leaf);
        box-shadow: 0 8px 20px rgba(47, 107, 79, .2);
    }
    .eyebrow { color: var(--leaf); font-size: .78rem; font-weight: 800; letter-spacing: .14em; }
    .hero { margin-bottom: .2rem; font-size: clamp(2rem, 4vw, 3.45rem); line-height: 1.05; letter-spacing: -.045em; }
    .subtitle { max-width: 780px; color: var(--muted); font-size: 1.04rem; line-height: 1.75; }
    .chips { display: flex; flex-wrap: wrap; gap: .55rem; margin: 1.15rem 0 2rem; }
    .chip {
        border: 1px solid #cfdcd3; border-radius: 999px; padding: .35rem .7rem;
        background: rgba(255,255,255,.64); color: #395647; font-size: .78rem; font-weight: 650;
    }
    .section-label { font-size: .78rem; font-weight: 800; color: #4d5c54; letter-spacing: .07em; }
    div[data-testid="stTextArea"] textarea {
        min-height: 620px; border: 1px solid var(--line); border-radius: 14px;
        background: var(--paper); color: #202820;
        font-family: SimSun, "宋体", serif !important;
        font-size: 10.5pt !important; line-height: 1.72 !important;
        box-shadow: 0 18px 48px rgba(27, 49, 36, .07);
    }
    div[data-testid="stTextArea"] textarea:focus { border-color: #6b9a7e; box-shadow: 0 0 0 3px rgba(73,128,94,.12); }
    [data-testid="stFileUploaderDropzone"] { min-height: 42px; padding: .3rem .55rem; border-radius: 10px; }
    [data-testid="stFileUploaderDropzoneInstructions"], [data-testid="stFileUploader"] small { display: none; }
    [data-testid="stFileUploaderDropzone"] button { margin: 0; min-height: 34px; }
    .st-key-preview_card {
        min-height: 620px; max-height: 620px; overflow: auto;
        padding: 2rem 2.35rem; border: 1px solid var(--line); border-radius: 14px;
        background: var(--paper); box-shadow: 0 18px 48px rgba(27, 49, 36, .07);
        font-family: SimSun, "宋体", serif; font-size: 10.5pt; line-height: 1.72;
    }
    .st-key-preview_card h1, .st-key-preview_card h2, .st-key-preview_card h3 { font-family: SimHei, "黑体", sans-serif; }
    .st-key-preview_card h1 { font-size: 1.65rem; }
    .st-key-preview_card h2 { font-size: 1.28rem; margin-top: 1.7rem; }
    .st-key-preview_card table { border-collapse: collapse; width: 100%; }
    .st-key-preview_card th, .st-key-preview_card td { border: 1px solid #d7dfda; padding: .45rem .6rem; }
    .st-key-preview_card blockquote { margin-left: 0; padding-left: 1rem; border-left: 3px solid #8ab39a; color: #59665f; }
    .status-note { color: var(--muted); font-size: .82rem; }
    div[data-testid="stButton"] button, div[data-testid="stDownloadButton"] button {
        border-radius: 10px; font-weight: 700;
    }
    div[data-testid="stDownloadButton"] button { background: var(--leaf); color: white; border-color: var(--leaf); }
    .footer-note { margin-top: 2.4rem; color: #7a857f; font-size: .78rem; text-align: center; }
    @media (max-width: 900px) {
      div[data-testid="stTextArea"] textarea, .st-key-preview_card { min-height: 430px; max-height: 430px; }
    }
</style>
""",
    unsafe_allow_html=True,
)

st.markdown('<div class="brand"><span class="brand-mark">墨</span> 墨转 MD2WORD</div>', unsafe_allow_html=True)
st.markdown('<div class="eyebrow">MARKDOWN → DOCX</div>', unsafe_allow_html=True)
st.markdown('<h1 class="hero">把 Markdown，变成真正好用的 Word</h1>', unsafe_allow_html=True)
st.markdown(
    '<p class="subtitle">专为中文写作与数学公式优化。粘贴内容、检查预览、下载文档——公式会保留为 Word 原生对象，之后仍可继续编辑。</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="chips"><span class="chip">✦ 可编辑公式</span><span class="chip">宋体 · 五号</span><span class="chip">支持表格与代码</span><span class="chip">不保存文稿</span></div>',
    unsafe_allow_html=True,
)

toolbar_left, toolbar_upload, toolbar_example, toolbar_clear = st.columns([5.5, 2.2, 1.2, 1])
with toolbar_left:
    st.markdown('<div class="section-label">工作台</div>', unsafe_allow_html=True)
with toolbar_upload:
    uploaded = st.file_uploader("导入 Markdown", type=["md", "markdown", "txt"], label_visibility="collapsed")
with toolbar_example:
    st.button("加载示例", use_container_width=True, on_click=load_example)
with toolbar_clear:
    st.button("清空", use_container_width=True, on_click=clear_editor)

if uploaded is not None:
    upload_key = f"{uploaded.name}:{uploaded.size}"
    if st.session_state.get("loaded_upload") != upload_key:
        try:
            st.session_state.md_content = uploaded.getvalue().decode("utf-8-sig")
            st.session_state.loaded_upload = upload_key
            st.session_state.suggested_name = safe_markdown_name(uploaded.name)
            st.rerun()
        except UnicodeDecodeError:
            st.error("文件不是 UTF-8 编码，请先另存为 UTF-8 后再上传。")

editor_col, preview_col = st.columns(2, gap="medium")
with editor_col:
    st.caption("MARKDOWN 编辑器  ·  输入区域同样使用宋体五号显示")
    markdown_text = st.text_area(
        "Markdown 内容",
        key="md_content",
        height=620,
        label_visibility="collapsed",
        placeholder="在这里粘贴 Markdown……",
    )
with preview_col:
    st.caption("实时预览  ·  公式使用 KaTeX 渲染")
    with st.container(border=False, key="preview_card"):
        if markdown_text.strip():
            st.markdown(markdown_text)
        else:
            st.markdown("*预览会显示在这里。*")

st.markdown("### 导出设置")
settings_left, settings_mid, action_col = st.columns([4.5, 2.5, 3])
with settings_left:
    filename = st.text_input(
        "文件名",
        value=st.session_state.get("suggested_name", "我的文档"),
        help="无需输入 .docx 后缀",
    )
with settings_mid:
    st.text_input("正文格式", value="正文 · 宋体 · 五号（10.5 pt）", disabled=True)
with action_col:
    st.write("")
    generate = st.button("生成 Word", type="primary", use_container_width=True, disabled=not markdown_text.strip())

content_hash = hashlib.sha256(markdown_text.encode("utf-8")).hexdigest()
if generate:
    with st.spinner("正在排版并写入 Word 原生公式……"):
        try:
            output_name = safe_markdown_name(filename) + ".docx"
            docx_bytes = convert_markdown_to_docx(markdown_text, title=safe_markdown_name(filename))
            report = inspect_docx(docx_bytes)
            st.session_state.download = {
                "name": output_name,
                "data": docx_bytes,
                "hash": content_hash,
                "report": report,
            }
        except ConversionError as exc:
            st.session_state.pop("download", None)
            st.error(str(exc))

download = st.session_state.get("download")
if download and download["hash"] == content_hash:
    report = download["report"]
    success_col, download_col = st.columns([7, 3])
    with success_col:
        formula_note = f"，检测到 {report['equation_count']} 个 Word 原生公式" if report["equation_count"] else ""
        st.success(f"文档已生成{formula_note}。普通段落已套用 Word“正文”样式（宋体五号）。")
    with download_col:
        st.download_button(
            "下载 .docx",
            data=download["data"],
            file_name=download["name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
elif download:
    st.info("内容已经修改，请重新生成 Word。")

with st.expander("公式写法与兼容性说明"):
    st.markdown(
        r"""
- 行内公式用 `$E=mc^2$`，独立公式用 `$$ ... $$`。
- 也支持 `\(...\)` 与 `\[...\]`，导出前会自动规范化。
- 导出结果使用 Word 原生 OMML 公式对象，可在 Microsoft Word 中双击编辑；复杂宏、自定义 LaTeX 宏包命令可能需要改写为标准 LaTeX。
- 当前版本不嵌入本机相对路径图片；网络图片能否写入取决于部署环境的网络访问权限。
        """
    )

st.markdown('<div class="footer-note">墨转 · 文稿仅在本次转换进程中临时处理，不写入数据库</div>', unsafe_allow_html=True)
