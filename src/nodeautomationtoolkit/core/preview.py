from __future__ import annotations

from pathlib import Path
from typing import Any

from .batch_types import WordDocumentBatch
from .table_types import DataTable
from .word_types import WordDocument, WordParagraphs, WordSaveResult


def format_live_preview(outputs: dict[str, Any], limit: int = 420) -> str:
    parts = []
    for name, value in outputs.items():
        parts.append(f"{name}: {_format_value(value)}")
    text = "\n".join(parts) if parts else "Виконано без вихідних даних"
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def _format_value(value: Any) -> str:
    if isinstance(value, DataTable):
        sample = " | ".join(str(item) for item in value.rows[0]) if value.rows else "порожньо"
        return f"{value.title} · {len(value.rows)} рядків\n{sample}"
    if isinstance(value, WordDocumentBatch):
        names = ", ".join(item.name for item in value.variants[:4])
        return (
            f"{len(value.variants)} документів · {len(value.operations)} операцій\n"
            f"{names}"
        )
    if isinstance(value, WordDocument):
        preview = _compact_text(value.text)
        return f"{value.file_name} · {value.paragraph_count} абзаців\n{preview}"
    if isinstance(value, WordParagraphs):
        return f"{len(value)} абзаців\n{_compact_text(value.text())}"
    if isinstance(value, WordSaveResult):
        return f"{value.file_name} · {value.paragraph_count} абзаців\n{value.path}"
    if isinstance(value, str):
        path = Path(value)
        if value and path.is_file():
            size_kb = path.stat().st_size / 1024
            content = _preview_file_content(path)
            suffix = f"\n{content}" if content else f"\n{path}"
            return f"{path.name} · {size_kb:.1f} KB{suffix}"
        return _compact_text(value)
    if isinstance(value, (list, tuple, set)):
        sample = ", ".join(str(item) for item in list(value)[:3])
        return f"{len(value)} елементів: {sample}"
    return repr(value)


def _compact_text(text: str, limit: int = 260) -> str:
    compact = " ".join(text.split())
    return compact if len(compact) <= limit else compact[: limit - 1].rstrip() + "…"


def _preview_file_content(path: Path) -> str:
    try:
        if path.suffix.casefold() == ".docx":
            from docx import Document

            document = Document(path)
            return _compact_text("\n".join(item.text for item in document.paragraphs))
        if path.suffix.casefold() in {".txt", ".md", ".json", ".csv", ".log"}:
            return _compact_text(path.read_text(encoding="utf-8", errors="replace"))
    except (OSError, ValueError):
        return ""
    return ""
