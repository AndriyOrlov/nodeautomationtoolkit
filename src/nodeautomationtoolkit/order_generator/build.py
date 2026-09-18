"""Збірка документа Word із проєкту наказу (етап «збірка»).

Верстка — за тими самими правилами, що й у витягах і примірниках
(PROJECT_RULES розд. 5), бо це той самий вид документа:

- Times New Roman 14 pt; блок виконавця — 8 pt звичайним;
- абзацний відступ пунктів 1.25 см;
- рівно 1 порожній абзац перед кожним пунктом і рівно 2 перед підписантом,
  жодного «висячого» порожнього абзацу в кінці;
- пункт разом зі своїм біографічним блоком не розривається між сторінками
  (KeepTogether), шапка зчеплена з першим пунктом, а останній пункт —
  з підписантом і засвідченням (KeepWithNext);
- нерозривні пробіли після коротких прийменників і в «ВОС - 0000000» уже
  стоять у тексті (`typography.apply_ukrainian_typography`).

Документ збирається python-docx, без Word COM: так збірка працює й на машині
без відкритого Word, і в тестах. Шаблон наказу («Шаблон наказу» в
налаштуваннях) не обов'язковий: без нього аркуш будується з нуля за геометрією
додатка 53 (А4, поля 2.0 / 1.0 / 1.0 / 1.0 см).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

from ..builtin_nodes.template_tags import expand_common_tags, signer_tags
from ..builtin_nodes.typography import ITEM_START_RE, ORDER_SIGNER_START_RE
from .compose import OrderDraft, numbered_lines

CONTENT_TAG = "{{зміст}}"
BODY_FONT = "Times New Roman"
BODY_SIZE = Pt(14)
EXECUTOR_SIZE = Pt(8)
FIRST_LINE_INDENT = Cm(1.25)
#: Біографічний блок пункту (р.н., освіта, у ЗС, РНОКПП, висновок про посаду)
#: стоїть не під самим пунктом, а з відступом у праву половину аркуша. Міряно
#: по зразках додатка 53: там лівий відступ таких абзаців 8.04–8.15 см.
BIO_LEFT_INDENT = Cm(8.0)


def _is_rank_subheading(line: str) -> bool:
    """««МАЙОР»» — підзаголовок групи в наказі про звання; він центрується."""
    clean = (line or "").replace(" ", " ").strip()
    return clean.startswith("«") and clean.endswith("»") and clean == clean.upper()


@dataclass
class OrderDocumentParts:
    """Що ще, крім пунктів, лягає в документ."""

    title: str = "НАКАЗ"
    subtitle: str = ""  # «по особовому складу»
    signer_position: str = ""
    signer_rank: str = ""
    signer_name: str = ""
    executor: str = ""


def _kind(line: str) -> str:
    """Рід абзацу: пункт, шапка чи продовження пункту (біографія, висновок)."""
    clean = (line or "").replace(" ", " ").strip()
    if not clean:
        return "blank"
    if clean.startswith("§"):
        return "heading"
    if ITEM_START_RE.match(clean):
        return "item"
    if ORDER_SIGNER_START_RE.match(clean):
        return "signer"
    if clean.endswith(":"):
        return "heading"
    # Підзаголовок групи в наказі про звання — саме звання ВЕЛИКИМИ в лапках
    # («МАЙОР»). Крапки в кінці він не має, тому окреме правило.
    if _is_rank_subheading(clean):
        return "heading"
    return "continuation"


def content_lines(draft: OrderDraft, parts: OrderDocumentParts | None = None) -> list[str]:
    """Рядки документа разом із порожніми: зміст, підписант, виконавець.

    Порожні рядки тут — це майбутні порожні абзаци: 1 перед пунктом,
    2 перед підписантом (PROJECT_RULES 5.3).
    """
    parts = parts or OrderDocumentParts()
    lines = list(numbered_lines(draft))
    signature = [
        part
        for part in (
            parts.signer_position,
            " ".join(piece for piece in (parts.signer_rank, parts.signer_name) if piece).strip(),
        )
        if part
    ]
    if signature:
        lines.extend(["", "", *signature])
    if parts.executor:
        lines.extend(["", parts.executor])
    while lines and not lines[-1].strip():
        lines.pop()
    return lines


def _style_paragraph(paragraph, line: str, executor: bool) -> None:
    """Геометрія абзацу за зразками додатка 53 і PROJECT_RULES розд. 5."""
    kind = _kind(line)
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.left_indent = Cm(0)
    if kind == "continuation":
        # Біографія й висновок про посаду — окремим блоком праворуч.
        fmt.alignment = WD_ALIGN_PARAGRAPH.LEFT
        fmt.left_indent = BIO_LEFT_INDENT
        fmt.first_line_indent = Cm(0)
    elif _is_rank_subheading(line):
        fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fmt.first_line_indent = Cm(0)
    elif kind in ("item", "heading"):
        fmt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        fmt.first_line_indent = FIRST_LINE_INDENT
    else:
        fmt.first_line_indent = Cm(0)
    for run in paragraph.runs:
        run.font.name = BODY_FONT
        run.font.size = EXECUTOR_SIZE if executor else BODY_SIZE
        run.font.italic = False


def _apply_keep_rules(paragraphs: list, lines: list[str]) -> None:
    """KeepTogether / KeepWithNext — правило PROJECT_RULES 5.4.

    Пункт тримає свій біографічний блок; шапка тримається за першим пунктом;
    останній пункт, підписант і засвідчення — один неподільний ланцюг.
    """
    kinds = [_kind(line) for line in lines]
    last_meaningful = max(
        (index for index, kind in enumerate(kinds) if kind != "blank"), default=-1
    )
    first_of_tail = None
    for index in range(last_meaningful, -1, -1):
        if kinds[index] == "item":
            first_of_tail = index
            break

    for index, paragraph in enumerate(paragraphs):
        fmt = paragraph.paragraph_format
        if kinds[index] == "blank":
            fmt.keep_together = False
        else:
            fmt.keep_together = True

        following = next(
            (j for j in range(index + 1, len(kinds)) if kinds[j] != "blank"), None
        )
        if following is None:
            keep_with_next = False
        elif kinds[index] == "heading":
            keep_with_next = True
        elif first_of_tail is not None and index >= first_of_tail:
            # Останній пункт → 2 ентери → підписант → засвідчення: єдиний ланцюг.
            keep_with_next = True
        else:
            keep_with_next = kinds[following] == "continuation"
        fmt.keep_with_next = keep_with_next
        # Порожній абзац усередині ланцюга не має його розривати, а порожній
        # між двома пунктами — навпаки, лишається місцем розриву сторінки.
        gap_end = following if following is not None else len(kinds)
        for blank_index in range(index + 1, gap_end):
            paragraphs[blank_index].paragraph_format.keep_with_next = keep_with_next


def _replace_tag_text(paragraph, values: dict[str, str]) -> None:
    text = paragraph.text
    replaced = text
    for tag, value in values.items():
        if tag in replaced:
            replaced = replaced.replace(tag, value)
    if replaced == text:
        return
    for run in paragraph.runs[1:]:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = replaced
    else:
        paragraph.add_run(replaced)


def _find_content_paragraph(document):
    for paragraph in document.paragraphs:
        if CONTENT_TAG in paragraph.text:
            return paragraph
    return None


def _new_document() -> Document:
    document = Document()
    section = document.sections[0]
    section.left_margin, section.right_margin = Cm(2.0), Cm(1.0)
    section.top_margin, section.bottom_margin = Cm(1.0), Cm(1.0)
    style = document.styles["Normal"]
    style.font.name = BODY_FONT
    style.font.size = BODY_SIZE
    return document


def build_order_document(
    draft: OrderDraft,
    output_path: str | Path,
    template_path: str | Path = "",
    parts: OrderDocumentParts | None = None,
) -> Path:
    """Збирає .docx наказу. Шаблон необов'язковий — без нього аркуш чистий."""
    parts = parts or OrderDocumentParts()
    lines = content_lines(draft, parts)
    executor_from = len(lines) - 1 if parts.executor and lines else len(lines)

    template = Path(template_path) if template_path else None
    if template and template.is_file():
        document = Document(str(template))
        anchor = _find_content_paragraph(document)
    else:
        document = _new_document()
        anchor = None
        if parts.title:
            heading = document.add_paragraph(parts.title)
            heading.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in heading.runs:
                run.font.name, run.font.size, run.font.bold = BODY_FONT, Pt(22), True
        if parts.subtitle:
            subtitle = document.add_paragraph(parts.subtitle)
            subtitle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in subtitle.runs:
                run.font.name, run.font.size = BODY_FONT, BODY_SIZE
        document.add_paragraph("")

    written = []
    for index, line in enumerate(lines):
        if anchor is not None:
            paragraph = anchor.insert_paragraph_before(line)
        else:
            paragraph = document.add_paragraph(line)
        _style_paragraph(paragraph, line, executor=index >= executor_from)
        written.append(paragraph)
    if anchor is not None:
        anchor._element.getparent().remove(anchor._element)

    _apply_keep_rules(written, lines)

    values = expand_common_tags(
        {
            "{{номер_наказу}}": draft.params.number,
            "{{дата_наказу}}": draft.params.date,
            "{{виконавець}}": parts.executor,
        }
    )
    values.update(signer_tags(parts.signer_position, parts.signer_rank, parts.signer_name))
    for paragraph in document.paragraphs:
        if "{{" in paragraph.text:
            _replace_tag_text(paragraph, values)

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))
    return output


def order_filename(draft: OrderDraft, prefix: str = "Наказ") -> str:
    """«Наказ № 525 від 17.09.2026.docx» — або з датою прогону, якщо реквізитів немає."""
    number = re.sub(r"[\\/:*?\"<>|]", "", str(draft.params.number or "")).strip()
    date = re.sub(r"[\\/:*?\"<>|]", "", str(draft.params.date or "")).strip()
    parts = [prefix]
    if number:
        parts.append(f"№ {number}")
    if date:
        parts.append(f"від {date}")
    return " ".join(parts) + ".docx"
