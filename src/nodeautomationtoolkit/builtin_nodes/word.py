from __future__ import annotations

from pathlib import Path

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.word_types import (
    WordDocument,
    WordParagraph,
    WordParagraphs,
    WordSaveResult,
)


def _document_class():
    from docx import Document

    return Document


def _validated_docx_path(path: str, *, must_exist: bool) -> Path:
    if not path.strip():
        raise ValueError("Шлях до DOCX не вказано")
    resolved = Path(path).expanduser()
    if resolved.suffix.casefold() != ".docx":
        resolved = resolved.with_suffix(".docx")
    if must_exist and not resolved.is_file():
        raise FileNotFoundError(f"DOCX не знайдено: {resolved}")
    return resolved


@node(
    name="Прочитати DOCX",
    category="Word",
    description="Читає локальний DOCX, його назву, текст і абзаци.",
    type_id="builtin.word.read_docx",
    outputs={
        "document": "WordDocument",
        "file_name": "str",
        "paragraphs": "WordParagraphs",
        "text": "str",
    },
)
def read_docx(path: str) -> dict:
    source_path = _validated_docx_path(path, must_exist=True)
    document = _document_class()(source_path)
    paragraph_items = tuple(
        WordParagraph(
            index=index,
            text=paragraph.text,
            style=paragraph.style.name if paragraph.style is not None else "",
            is_empty=not paragraph.text.strip(),
        )
        for index, paragraph in enumerate(document.paragraphs)
    )
    paragraphs = WordParagraphs(paragraph_items, str(source_path))
    text = paragraphs.text()
    word_document = WordDocument(
        path=str(source_path),
        file_name=source_path.name,
        paragraphs=paragraphs,
        text=text,
    )
    return {
        "document": word_document,
        "file_name": word_document.file_name,
        "paragraphs": paragraphs,
        "text": text,
    }


@node(
    name="Абзаци документа",
    category="Word",
    description="Повертає структуровані абзаци прочитаного Word-документа.",
    type_id="builtin.word.document_paragraphs",
)
def document_paragraphs(document: WordDocument) -> WordParagraphs:
    return document.paragraphs


@node(
    name="Текст документа",
    category="Word",
    description="Повертає суцільний текст прочитаного Word-документа.",
    type_id="builtin.word.document_text",
)
def document_text(document: WordDocument, include_empty: bool = False) -> str:
    if include_empty:
        return document.text
    return "\n".join(item.text for item in document.paragraphs if not item.is_empty)


@node(
    name="Знайти абзаци",
    category="Word",
    description="Вибирає абзаци, у тексті яких є вказаний збіг.",
    type_id="builtin.word.filter_paragraphs",
)
def filter_paragraphs(
    paragraphs: WordParagraphs,
    contains: str,
    ignore_case: bool = True,
) -> WordParagraphs:
    if not contains:
        return paragraphs
    needle = contains.casefold() if ignore_case else contains
    selected = []
    for paragraph in paragraphs:
        candidate = paragraph.text.casefold() if ignore_case else paragraph.text
        if needle in candidate:
            selected.append(paragraph)
    return WordParagraphs(tuple(selected), paragraphs.source_path)


@node(
    name="Абзаци в текст",
    category="Word",
    description="Об'єднує вибрані Word-абзаци у текст.",
    type_id="builtin.word.paragraphs_to_text",
)
def paragraphs_to_text(paragraphs: WordParagraphs, separator: str = "\n") -> str:
    return paragraphs.text(separator)


@node(
    name="Створити DOCX",
    category="Word",
    description="Створює новий Word-документ із тексту й зберігає його локально.",
    type_id="builtin.word.create_docx",
    execution_inputs=("exec",),
    execution_outputs=("then",),
)
def create_docx(text: str, output_path: str, title: str = "") -> WordSaveResult:
    target_path = _validated_docx_path(output_path, must_exist=False)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    document = _document_class()()
    if title.strip():
        document.add_heading(title.strip(), level=1)
    lines = text.splitlines() or [text]
    for line in lines:
        paragraph = document.add_paragraph(line)
        paragraph.paragraph_format.keep_together = True
    document.save(target_path)
    return WordSaveResult(
        path=str(target_path),
        paragraph_count=len(document.paragraphs),
    )


@node(
    name="Зберегти вибрані абзаци",
    category="Word",
    description=(
        "Створює копію вихідного DOCX, залишає вибрані абзаци та зберігає форматування."
    ),
    type_id="builtin.word.save_selected_paragraphs",
    execution_inputs=("exec",),
    execution_outputs=("then",),
)
def save_selected_paragraphs(
    document: WordDocument,
    paragraphs: WordParagraphs,
    output_path: str,
    keep_together: bool = True,
) -> WordSaveResult:
    target_path = _validated_docx_path(output_path, must_exist=False)
    target_path.parent.mkdir(parents=True, exist_ok=True)
    source = _document_class()(document.path)
    selected_indices = {item.index for item in paragraphs}

    for index, paragraph in reversed(list(enumerate(source.paragraphs))):
        if index not in selected_indices:
            element = paragraph._element
            element.getparent().remove(element)
        elif keep_together:
            paragraph.paragraph_format.keep_together = True

    for table in list(source.tables):
        element = table._element
        element.getparent().remove(element)

    source.save(target_path)
    return WordSaveResult(
        path=str(target_path),
        paragraph_count=len(selected_indices),
        message="Вибрані абзаци збережено",
    )
