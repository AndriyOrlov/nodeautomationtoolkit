"""Модуль аналізу та порівняння документів Word (.docx) (Compare Mode).

Забезпечує 100% конфіденційність: звіт для чату AI містить виключно
технічні правила, параметри форматування (відступи, вирівнювання, ентери, шифри),
БЕЗ жодного фрагмента персональних даних чи тексту наказів.
"""

from __future__ import annotations

import difflib
import os
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NAMESPACES = {"w": _W_NS}

_ALIGN_MAP = {
    "center": "По середині (Center)",
    "both": "По ширині (Justify)",
    "left": "По лівому краю (Left)",
    "right": "По правому краю (Right)",
}

_ORDER_SIGNER_START_RE = re.compile(
    r"^\s*(?:т\.?\s*в\.?\s*о\.?|тимчасово\s+виконуюч(?:ий|а)?|"
    r"командувач|командир|начальник|заступник\s+командувача)\b",
    re.IGNORECASE | re.UNICODE,
)


@dataclass
class DocParagraph:
    """Структурований опис абзацу документа."""

    index: int
    text: str
    alignment: str = "По лівому краю (Left)"
    raw_align: str = "left"
    first_line_indent_pt: float = 0.0
    left_indent_pt: float = 0.0
    font_size_pt: float = 14.0
    is_bold: bool = False
    is_blank: bool = False
    is_item: bool = False
    item_label: str = ""
    is_bio: bool = False
    is_signer: bool = False
    is_certifier: bool = False
    has_non_breaking_spaces: bool = False


@dataclass
class DiffDiscrepancy:
    """Опис окремої знайденої розбіжності правил (без витоку тексту)."""

    item_label: str
    issue_type: str
    description: str
    expected_rule: str = ""
    actual_rule: str = ""
    fix_suggestion: str = ""


@dataclass
class CompareResult:
    """Результат порівняння двох документів."""

    ref_path: str
    gen_path: str
    mode: str
    is_identical: bool = True
    discrepancies: list[DiffDiscrepancy] = field(default_factory=list)
    side_by_side_rows: list[dict] = field(default_factory=list)
    summary_text: str = ""
    ai_chat_report: str = ""


def _parse_docx_paragraphs(docx_path: str | Path) -> list[DocParagraph]:
    """Швидкий XML-парсер структури абзаців із DOCX файлу."""
    path = Path(docx_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл не знайдено: {path}")

    paragraphs: list[DocParagraph] = []
    with zipfile.ZipFile(path, "r") as zf:
        if "word/document.xml" not in zf.namelist():
            return paragraphs
        xml_content = zf.read("word/document.xml")
        root = ET.fromstring(xml_content)

    body = root.find("w:body", _NAMESPACES)
    if body is None:
        return paragraphs

    p_idx = 0
    for elem in body:
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag
        if tag == "p":
            p_idx += 1
            # 1. Текст
            text_parts = []
            font_size = 14.0
            is_bold = False
            for r in elem.findall("w:r", _NAMESPACES):
                r_pr = r.find("w:rPr", _NAMESPACES)
                if r_pr is not None:
                    sz_elem = r_pr.find("w:sz", _NAMESPACES)
                    if sz_elem is not None:
                        val = sz_elem.attrib.get(f"{{{_W_NS}}}val") or sz_elem.attrib.get("val")
                        if val and val.isdigit():
                            font_size = float(val) / 2.0
                    b_elem = r_pr.find("w:b", _NAMESPACES)
                    if b_elem is not None:
                        is_bold = True
                for t in r.findall("w:t", _NAMESPACES):
                    if t.text:
                        text_parts.append(t.text)
                for tab in r.findall("w:tab", _NAMESPACES):
                    text_parts.append("\t")

            full_text = "".join(text_parts).rstrip("\r\n")
            clean_text = full_text.strip()
            is_blank = not bool(clean_text)

            # 2. Властивості абзацу
            raw_align = "left"
            first_indent = 0.0
            left_indent = 0.0
            p_pr = elem.find("w:pPr", _NAMESPACES)
            if p_pr is not None:
                jc = p_pr.find("w:jc", _NAMESPACES)
                if jc is not None:
                    raw_align = jc.attrib.get(f"{{{_W_NS}}}val") or jc.attrib.get("val") or "left"

                ind = p_pr.find("w:ind", _NAMESPACES)
                if ind is not None:
                    f_val = ind.attrib.get(f"{{{_W_NS}}}firstLine") or ind.attrib.get("firstLine")
                    if f_val and f_val.lstrip("-").isdigit():
                        first_indent = round(float(f_val) / 20.0, 2)
                    l_val = ind.attrib.get(f"{{{_W_NS}}}left") or ind.attrib.get("left")
                    if l_val and l_val.lstrip("-").isdigit():
                        left_indent = round(float(l_val) / 20.0, 2)

                p_rpr = p_pr.find("w:rPr", _NAMESPACES)
                if p_rpr is not None:
                    sz_elem = p_rpr.find("w:sz", _NAMESPACES)
                    if sz_elem is not None:
                        val = sz_elem.attrib.get(f"{{{_W_NS}}}val") or sz_elem.attrib.get("val")
                        if val and val.isdigit():
                            font_size = float(val) / 2.0

            # 3. Семантичні прапорці
            low_text = clean_text.lower()
            item_match = re.match(r"^(\d{1,3}(?:\.\d{1,3})*[\.\)])\s+", clean_text)
            is_item = bool(item_match) and not clean_text.startswith("§")
            item_label = item_match.group(1) if item_match else ""

            is_bio = bool(
                ("р.н." in low_text or "р. н." in low_text or "року народження" in low_text or "освіта:" in low_text or "у зс із" in low_text or "іпн" in low_text or "рнокпп" in low_text or "вос-" in low_text)
                and not is_item
                and not clean_text.startswith("§")
            )
            is_signer = bool(_ORDER_SIGNER_START_RE.match(clean_text))
            is_certifier = bool("згідно з оригіналом" in low_text or "т.в.о. начальника штабу" in low_text)
            has_nbsp = "\u00A0" in full_text

            paragraphs.append(
                DocParagraph(
                    index=p_idx,
                    text=full_text,
                    alignment=_ALIGN_MAP.get(raw_align, "По лівому краю (Left)"),
                    raw_align=raw_align,
                    first_line_indent_pt=first_indent,
                    left_indent_pt=left_indent,
                    font_size_pt=font_size,
                    is_bold=is_bold,
                    is_blank=is_blank,
                    is_item=is_item,
                    item_label=item_label,
                    is_bio=is_bio,
                    is_signer=is_signer,
                    is_certifier=is_certifier,
                    has_non_breaking_spaces=has_nbsp,
                )
            )

    return paragraphs


def _normalize_text_for_compare(t: str) -> str:
    """Нормалізує текст для точного порівняння структури (ігноруючи незначні варіації лапок/тире/пробілів)."""
    if not t:
        return ""
    # Нерозривні та вузькі пробіли
    t = t.replace("\u00a0", " ").replace("\u202f", " ").replace("\t", " ")
    # Лапки
    t = t.replace("«", '"').replace("»", '"').replace("“", '"').replace("”", '"').replace("„", '"')
    # Апострофи
    t = t.replace("’", "'").replace("ʼ", "'").replace("`", "'")
    # Тире / дефіси
    t = t.replace("–", "-").replace("—", "-").replace("―", "-")
    # Згортання пробілів
    return " ".join(t.split())


def compare_docx_documents(
    reference_path: str | Path,
    generated_path: str | Path,
    mode: str = "extract",
) -> CompareResult:
    """Попарно порівнює еталонний та згенерований DOCX документи."""
    ref_paras = _parse_docx_paragraphs(reference_path)
    gen_paras = _parse_docx_paragraphs(generated_path)

    discrepancies: list[DiffDiscrepancy] = []

    # 1. Аналіз послідовності структури через SequenceMatcher з нормалізацією пробілів
    ref_clean_lines = [_normalize_text_for_compare(p.text) for p in ref_paras]
    gen_clean_lines = [_normalize_text_for_compare(p.text) for p in gen_paras]

    matcher = difflib.SequenceMatcher(None, ref_clean_lines, gen_clean_lines)
    side_by_side_rows = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            for r_idx, g_idx in zip(range(i1, i2), range(j1, j2)):
                rp = ref_paras[r_idx]
                gp = gen_paras[g_idx]
                side_by_side_rows.append(
                    {
                        "status": "EQUAL",
                        "ref_line": rp.text,
                        "gen_line": gp.text,
                        "ref_p": rp,
                        "gen_p": gp,
                    }
                )
                # Перевіряємо форматування однакових рядків
                _check_formatting_discrepancy(rp, gp, discrepancies)

        elif tag == "replace":
            r_slice = ref_paras[i1:i2]
            g_slice = gen_paras[j1:j2]
            max_len = max(len(r_slice), len(g_slice))
            for k in range(max_len):
                rp = r_slice[k] if k < len(r_slice) else None
                gp = g_slice[k] if k < len(g_slice) else None

                r_text = rp.text if rp else ""
                g_text = gp.text if gp else ""

                if rp and gp and _normalize_text_for_compare(rp.text) == _normalize_text_for_compare(gp.text):
                    status = "EQUAL"
                    side_by_side_rows.append(
                        {
                            "status": status,
                            "ref_line": r_text,
                            "gen_line": g_text,
                            "ref_p": rp,
                            "gen_p": gp,
                        }
                    )
                    _check_formatting_discrepancy(rp, gp, discrepancies)
                    continue

                status = "MODIFIED"
                side_by_side_rows.append(
                    {
                        "status": status,
                        "ref_line": r_text,
                        "gen_line": g_text,
                        "ref_p": rp,
                        "gen_p": gp,
                    }
                )

                if rp and gp:
                    label = _get_semantic_label(rp, gp)
                    _analyze_text_and_format_discrepancy(label, rp, gp, discrepancies)
                elif rp and not gp:
                    label = _get_semantic_label(rp, None)
                    discrepancies.append(
                        DiffDiscrepancy(
                            item_label=label,
                            issue_type="Пропущений пункт / абзац",
                            description="В еталоні присутній структурний блок, якого немає у згенерованому витязі.",
                            expected_rule="Пункт включено до витягу",
                            actual_rule="Пункт відсутній у згенерованому файлі",
                            fix_suggestion="Перевірити мапу маршрутизації та умови включення пункту.",
                        )
                    )
                elif gp and not rp:
                    label = _get_semantic_label(None, gp)
                    discrepancies.append(
                        DiffDiscrepancy(
                            item_label=label,
                            issue_type="Зайвий пункт / абзац",
                            description="У згенерованому файлі з'явився блок, якого немає в еталонному документі.",
                            expected_rule="Пункт відсутній у витягу цієї частини",
                            actual_rule="Пункт помилково потрапив до витягу",
                            fix_suggestion="Перевірити умови фільтрації адресатів для даного пункту.",
                        )
                    )

        elif tag == "delete":
            for r_idx in range(i1, i2):
                rp = ref_paras[r_idx]
                side_by_side_rows.append(
                    {
                        "status": "DELETED",
                        "ref_line": rp.text,
                        "gen_line": "",
                        "ref_p": rp,
                        "gen_p": None,
                    }
                )
                if not rp.is_blank:
                    label = _get_semantic_label(rp, None)
                    discrepancies.append(
                        DiffDiscrepancy(
                            item_label=label,
                            issue_type="Пропущений структурний блок",
                            description="Абзац із еталона не потрапив у згенерований документ.",
                            expected_rule="Блок присутній",
                            actual_rule="Блок відсутній",
                            fix_suggestion="Перевірити правила формування витягу для цієї ВЧ.",
                        )
                    )

        elif tag == "insert":
            for g_idx in range(j1, j2):
                gp = gen_paras[g_idx]
                side_by_side_rows.append(
                    {
                        "status": "INSERTED",
                        "ref_line": "",
                        "gen_line": gp.text,
                        "ref_p": None,
                        "gen_p": gp,
                    }
                )
                if not gp.is_blank:
                    label = _get_semantic_label(None, gp)
                    discrepancies.append(
                        DiffDiscrepancy(
                            item_label=label,
                            issue_type="Зайвий блок у документі",
                            description="Згенеровано додатковий блок, відсутній у ручному еталоні.",
                            expected_rule="Блок відсутній",
                            actual_rule="Блок згенеровано",
                            fix_suggestion="Перевірити видалення зайвих службових рядків.",
                        )
                    )

    # 2. Перевірка правил ентерів (1 порожній рядок перед пунктом, 2 перед підписантом)
    _check_blank_line_rules(gen_paras, discrepancies)

    is_identical = len(discrepancies) == 0

    # 3. Формування зведення та конфіденційного звіту для чату AI
    summary_text = _build_summary_text(discrepancies)
    ai_chat_report = generate_ai_chat_report(
        discrepancies=discrepancies,
        mode=mode,
    )

    return CompareResult(
        ref_path=str(reference_path),
        gen_path=str(generated_path),
        mode=mode,
        is_identical=is_identical,
        discrepancies=discrepancies,
        side_by_side_rows=side_by_side_rows,
        summary_text=summary_text,
        ai_chat_report=ai_chat_report,
    )


def _get_semantic_label(rp: DocParagraph | None, gp: DocParagraph | None) -> str:
    """Визначає безпечну назву блоку без персональних даних."""
    p = rp or gp
    if not p:
        return "Абзац документа"
    if p.is_item and p.item_label:
        return f"Пункт {p.item_label}"
    if p.is_bio:
        return "Біографічний блок (р.н. / освіта / служба)"
    if p.is_signer:
        return "Блок підписанта наказу"
    if p.is_certifier:
        return "Блок завірителя («Згідно з оригіналом»)"
    return f"Абзац № {p.index}"


def _check_formatting_discrepancy(
    rp: DocParagraph,
    gp: DocParagraph,
    discrepancies: list[DiffDiscrepancy],
):
    """Перевіряє форматування для ідентичних за текстом рядків (без витоку тексту)."""
    if rp.is_blank and gp.is_blank:
        return

    label = _get_semantic_label(rp, gp)

    # Вирівнювання
    if rp.raw_align != gp.raw_align and not rp.is_blank:
        discrepancies.append(
            DiffDiscrepancy(
                item_label=label,
                issue_type="Вирівнювання (Alignment)",
                description=f"Не співпадає вирівнювання: в еталоні '{rp.alignment}', а згенеровано '{gp.alignment}'.",
                expected_rule=rp.alignment,
                actual_rule=gp.alignment,
                fix_suggestion=f"Встановити Alignment = {rp.raw_align} для {label}.",
            )
        )

    # Абзацний відступ (First Line Indent)
    if rp.is_item and abs(rp.first_line_indent_pt - gp.first_line_indent_pt) > 5.0:
        discrepancies.append(
            DiffDiscrepancy(
                item_label=label,
                issue_type="Абзацний відступ (Indent)",
                description=f"Відступ першого рядка відрізняється: очікувалось {rp.first_line_indent_pt} pt (1.25 см), отримано {gp.first_line_indent_pt} pt.",
                expected_rule=f"FirstLineIndent = {rp.first_line_indent_pt} pt (1.25 см)",
                actual_rule=f"FirstLineIndent = {gp.first_line_indent_pt} pt",
                fix_suggestion="Встановити FirstLineIndent = 35.45 pt (1.25 см) для пунктів наказу.",
            )
        )

    # Розмір шрифту
    if abs(rp.font_size_pt - gp.font_size_pt) > 0.5 and not rp.is_blank:
        discrepancies.append(
            DiffDiscrepancy(
                item_label=label,
                issue_type="Розмір шрифту (Font Size)",
                description=f"Розмір шрифту відрізняється: очікувалось {rp.font_size_pt} pt, отримано {gp.font_size_pt} pt.",
                expected_rule=f"{rp.font_size_pt} pt",
                actual_rule=f"{gp.font_size_pt} pt",
                fix_suggestion="Переконатися, що розмір шрифту встановлено рівно 14 pt.",
            )
        )


def _analyze_text_and_format_discrepancy(
    label: str,
    rp: DocParagraph,
    gp: DocParagraph,
    discrepancies: list[DiffDiscrepancy],
):
    """Детальний аналіз відмінностей правил (шифри, дублювання фраз, без цитування тексту наказу)."""
    # Перевірка на дублікати "військової частини"
    if "військової частини військової частини" in gp.text.lower() or "в/ч в/ч" in gp.text.lower():
        discrepancies.append(
            DiffDiscrepancy(
                item_label=label,
                issue_type="Повтор назви / шифру ВЧ",
                description="Виявлено подвійне повторення службової фрази 'військової частини' або шифру підряд.",
                expected_rule="Одинарна фраза 'військової частини А****'",
                actual_rule="Подвоєна фраза 'військової частини військової частини А****'",
                fix_suggestion="Застосувати clean_duplicated_units() для автоматичного згортання дублювання.",
            )
        )
        return

    # Перевірка розбіжностей шифрів ВЧ
    ref_ciphers = set(re.findall(r"\b[АA]\s*\d{4}\b", rp.text))
    gen_ciphers = set(re.findall(r"\b[АA]\s*\d{4}\b", gp.text))
    if ref_ciphers != gen_ciphers:
        discrepancies.append(
            DiffDiscrepancy(
                item_label=label,
                issue_type="Підстановка шифру ВЧ (Cipher)",
                description="Шифр закритої військової частини у згенерованому документі відрізняється від еталона.",
                expected_rule="Підстановка згідно зі словником Excel",
                actual_rule="Підставлено інший або нерозпізнаний шифр",
                fix_suggestion="Перевірити відповідність назви ВЧ та шифру в таблиці Excel / recipient_mapping.",
            )
        )
        return

    # Загальна структурна розбіжність
    discrepancies.append(
        DiffDiscrepancy(
            item_label=label,
            issue_type="Розбіжність структури або формулювання",
            description="Виявлено розбіжність у структурі пункту між еталоном та генератором.",
            expected_rule="Збереження оригінального тексту наказу з підстановкою шифру",
            actual_rule="Текст пункту зазнав видозмін",
            fix_suggestion="Перевірити збереження цілісності рядків пункту під час генерації.",
        )
    )


def _check_blank_line_rules(gen_paras: list[DocParagraph], discrepancies: list[DiffDiscrepancy]):
    """Перевіряє дотримання правил порожніх рядків (1 перед пунктом, 2 перед підписантом)."""
    for idx, p in enumerate(gen_paras):
        if p.is_item and idx > 0:
            prev_p = gen_paras[idx - 1]
            if not prev_p.is_blank:
                discrepancies.append(
                    DiffDiscrepancy(
                        item_label=f"Пункт {p.item_label}",
                        issue_type="Відсутній ентер перед пунктом",
                        description=f"Перед пунктом {p.item_label} відсутній обов'язковий порожній рядок (1 строковий абзац).",
                        expected_rule="Рівно 1 порожній рядок (1 Enter) перед початком пункту",
                        actual_rule="0 порожніх рядків (пункт зливається з попереднім)",
                        fix_suggestion="Забезпечити виклик ensure_blank_line_before_items() перед формуванням документа.",
                    )
                )

        if p.is_signer and idx > 1:
            prev1 = gen_paras[idx - 1]
            prev2 = gen_paras[idx - 2]
            if not (prev1.is_blank and prev2.is_blank):
                discrepancies.append(
                    DiffDiscrepancy(
                        item_label="Блок підписанта наказу",
                        issue_type="Відступ перед підписантом",
                        description="Перед блоком підписанта має бути рівно 2 порожні рядки (2 строкові абзаци).",
                        expected_rule="Рівно 2 порожні рядки перед підписантом",
                        actual_rule="Менше 2 порожніх рядків",
                        fix_suggestion="Забезпечити наявність рівно 2 порожніх рядків (нормалізація відступу перед підписантом).",
                    )
                )


def _build_summary_text(discrepancies: list[DiffDiscrepancy]) -> str:
    """Формує коротке резюме для плашки статусу."""
    if not discrepancies:
        return "✅ Документи повністю відповідають усім правилам! Розбіжностей не виявлено."
    counts: dict[str, int] = {}
    for d in discrepancies:
        counts[d.issue_type] = counts.get(d.issue_type, 0) + 1
    details = ", ".join(f"{k}: {v}" for k, v in counts.items())
    return f"⚠️ Знайдено розбіжностей: {len(discrepancies)} ({details})"


def generate_ai_chat_report(
    discrepancies: list[DiffDiscrepancy],
    mode: str = "extract",
) -> str:
    """Генерує стислий 100% конфіденційний звіт правил для AI (без тексту наказу).

    Розбіжності групуються за типом помилки (issue_type), щоб звіт лишався
    компактним навіть при десятках однотипних порушень в одному наказі.
    """
    if not discrepancies:
        return (
            f"### 📋 Результат перевірки Compare Mode\n\n"
            f"- **Режим:** `{mode}`\n\n"
            f"**Розбіжностей не виявлено.** Згенерований документ повністю відповідає еталонним правилам!"
        )

    groups: dict[str, list[DiffDiscrepancy]] = {}
    for d in discrepancies:
        groups.setdefault(d.issue_type, []).append(d)

    lines = [
        "### 🔍 Звіт розбіжностей правил Compare Mode (Конфіденційно: тільки правила та параметри)\n",
        f"- **Режим:** `{mode}`",
        f"- **Кількість виявлених порушень:** `{len(discrepancies)}` ({len(groups)} типів)\n",
        "---",
    ]

    for issue_type, group in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        labels = list(dict.fromkeys(d.item_label for d in group))
        labels_preview = ", ".join(labels[:6])
        if len(labels) > 6:
            labels_preview += f" … ще {len(labels) - 6}"
        sample = group[0]
        lines.append(f"\n**{issue_type}** — `{len(group)}` шт. ({labels_preview})")
        if sample.expected_rule:
            lines.append(f"- Очікуване правило: `{sample.expected_rule}`")
        if sample.actual_rule:
            lines.append(f"- Фактично у генераторі: `{sample.actual_rule}`")
        if sample.fix_suggestion:
            lines.append(f"- 💡 Що виправити в коді: {sample.fix_suggestion}")

    lines.append("\n---\n**Запит до AI:** Будь ласка, виправ зазначені правила та логіку обробки в коді програми.")
    return "\n".join(lines)
