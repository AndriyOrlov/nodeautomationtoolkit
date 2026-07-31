from docx import Document

from nodeautomationtoolkit.core.preview import format_live_preview


def test_file_preview_shows_docx_text(tmp_path):
    path = tmp_path / "order.docx"
    document = Document()
    document.add_paragraph("Перший видимий абзац наказу")
    document.save(path)

    preview = format_live_preview({"path": str(path)})

    assert "order.docx" in preview
    assert "Перший видимий абзац наказу" in preview
