from __future__ import annotations

from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_app_loads_and_generates_download() -> None:
    app_path = Path(__file__).resolve().parents[1] / "app.py"
    app = AppTest.from_file(app_path, default_timeout=30).run()
    assert not app.exception
    assert len(app.text_area) == 1
    assert [button.label for button in app.button] == ["加载示例", "清空", "生成 Word"]

    app.button[2].click().run()
    assert not app.exception
    assert len(app.get("download_button")) == 1
    assert "Word“正文”样式" in app.success[0].value
