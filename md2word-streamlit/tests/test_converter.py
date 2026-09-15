from __future__ import annotations

import io
import zipfile

from docx import Document

from converter import convert_markdown_to_docx, inspect_docx


def test_exports_editable_omml_and_five_size_simsun() -> None:
    data = convert_markdown_to_docx(
        r"""# 测试文档

这是正文，行内公式 $E=mc^2$。

$$
\frac{-b\pm\sqrt{b^2-4ac}}{2a}
$$

$$
A=\begin{bmatrix}1&2\\3&4\end{bmatrix}
$$
""",
        title="自动测试",
    )

    report = inspect_docx(data)
    assert report["equation_count"] >= 2
    assert report["has_simsun"] is True
    assert report["has_five_size"] is True

    document = Document(io.BytesIO(data))
    body_paragraph = next(paragraph for paragraph in document.paragraphs if "这是正文" in paragraph.text)
    assert body_paragraph.style.name == "正文"
    body_run = next(run for run in body_paragraph.runs if "这是正文" in run.text)
    assert body_run.font.size.pt == 10.5

    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        xml = archive.read("word/document.xml")
        settings = archive.read("word/settings.xml")
        assert b"<m:oMath" in xml
        assert b'w:percent="100"' in settings
        assert xml.find(b"<m:count") < xml.find(b"<m:mcJc")


def test_supports_bracket_math_delimiters() -> None:
    data = convert_markdown_to_docx(r"结果为 \[x=\frac{1}{2}\]。")
    assert inspect_docx(data)["equation_count"] >= 1
