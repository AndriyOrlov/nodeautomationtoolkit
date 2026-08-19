"""Окремий ізольований модуль генерації повідомлень (Message Order Module).

Відповідає за:
- Формування тексту закритого наказу/повідомлення зі змістом,
- Автозаміну відкритих найменувань на закриті шифри ВЧ/корпусів,
- Складання списку розсилки (адресатів),
- Побудову документа повідомлення зі змістом та супровідного листа.

Повна ізоляція: зміни у цьому модулі жодним чином не впливають на модуль витягів (extracts).
"""

from __future__ import annotations

import os
import re
from pathlib import Path

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.table_types import DataTable
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    _build_unit_fuzzy_pattern,
    _UNIT_PHRASE_REPLACEMENTS,
    _fix_military_typos,
    _match_case,
    _short_closed_code,
    _format_full_closed_unit_text,
    _ORDER_SIGNER_START_RE,
)


def _apply_custom_rules(text: str, rules_input: str | list | dict | None) -> tuple[str, int]:
    """Застосовує користувацькі правила замін (case-insensitive)."""
    if not text or not rules_input:
        return text, 0

    rule_dict = {}
    if isinstance(rules_input, str):
        for line in rules_input.splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "->" in line:
                src, dst = line.split("->", 1)
                rule_dict[src.strip()] = dst.strip()
            elif "=" in line:
                src, dst = line.split("=", 1)
                rule_dict[src.strip()] = dst.strip()
    elif isinstance(rules_input, list):
        for item in rules_input:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                rule_dict[str(item[0]).strip()] = str(item[1]).strip()
            elif isinstance(item, dict):
                src = str(item.get("from") or item.get("src") or "").strip()
                dst = str(item.get("to") or item.get("dst") or "").strip()
                if src:
                    rule_dict[src] = dst
    elif isinstance(rules_input, dict):
        for k, v in rules_input.items():
            rule_dict[str(k).strip()] = str(v).strip()

    count = 0
    res_text = text
    for old_val, new_val in rule_dict.items():
        if not old_val or not isinstance(new_val, str):
            continue
        clean_old = _fix_military_typos(str(old_val))
        pattern = re.compile(re.escape(clean_old), re.IGNORECASE)
        if pattern.search(res_text):
            matches = pattern.findall(res_text)
            count += len(matches)
            res_text = pattern.sub(lambda m: _match_case(m.group(0), new_val), res_text)
        else:
            orig_pattern = re.compile(re.escape(str(old_val)), re.IGNORECASE)
            if orig_pattern.search(res_text):
                matches = orig_pattern.findall(res_text)
                count += len(matches)
                res_text = orig_pattern.sub(lambda m: _match_case(m.group(0), new_val), res_text)

    return res_text, count


@node(
    name="Формування наказу по рішенню (Повідомлення)",
    category="Наказ",
    description="Формує повний закритий наказ для режиму повідомлень/супроводу.",
    type_id="builtin.order.generate_decision_order",
    execution_inputs=("exec",),
    execution_outputs=("then",),
    outputs={
        "decision_text": "str",
        "table": "DataTable",
        "replaced_count": "int",
        "summary": "str",
    },
)
def generate_decision_order(
    text: str = "",
    mapping: dict | None = None,
    new_header: str = "",
    fuzzy_match: bool = True,
    rules: str | list | dict | None = None,
) -> dict:
    """Генерує закритий текст для режиму повідомлення."""
    if not text.strip():
        return {
            "decision_text": "",
            "table": DataTable(("Відкрита назва", "Закритий код ВЧ", "Армійський корпус", "Форматована заміна у тексті"), ()),
            "replaced_count": 0,
            "summary": "Порожній текст наказу",
        }

    text = _fix_military_typos(text)
    mapping_raw = mapping or {}
    mapping_dict = {}
    for k, v in mapping_raw.items():
        clean_k = _fix_military_typos(str(k))
        if isinstance(v, dict):
            clean_v = dict(v)
            if "open_name" in clean_v:
                clean_v["open_name"] = _fix_military_typos(str(clean_v["open_name"]))
            mapping_dict[clean_k] = clean_v
        else:
            mapping_dict[clean_k] = v

    lines = [line.rstrip() for line in text.splitlines()]

    # 1. Знаходимо початок змістовної частини (після шапки наказу)
    content_start_idx = 0
    for idx, line in enumerate(lines):
        clean = line.strip()
        if (
            clean.startswith("§")
            or re.match(r"^\d+[\.\)]", clean)
            or "НАКАЗУЮ" in clean.upper()
            or "ПРИЗНАЧИТИ" in clean.upper()
            or "НАПРАВИТИ" in clean.upper()
            or "ВІДРЯДИТИ" in clean.upper()
            or "ЗВІЛЬНИТИ" in clean.upper()
            or "ВІЙСЬКОВОСЛУЖБОВЦІВ" in clean.upper()
        ):
            content_start_idx = idx
            break

    body_lines = lines[content_start_idx:]
    body_text = "\n".join(body_lines)

    replaced_count = 0
    report_rows = []

    # 2. Створюємо список патернів з урахуванням відмінків
    patterns_to_apply = []
    for open_name, mapped_val in mapping_dict.items():
        if not open_name or not str(open_name).strip():
            continue
        closed_code = _format_full_closed_unit_text(mapped_val, mapping_dict)
        if isinstance(mapped_val, dict):
            raw_cipher = str(mapped_val.get("cipher") or "")
            corps_info = str(mapped_val.get("corps") or "")
            abbreviation = str(mapped_val.get("abbreviation") or "").strip()
        else:
            raw_cipher = str(mapped_val)
            corps_info = ""
            abbreviation = ""

        pattern = _build_unit_fuzzy_pattern(open_name) if fuzzy_match else re.compile(rf"\b{re.escape(open_name)}\b", re.IGNORECASE)
        patterns_to_apply.append((len(open_name), pattern, closed_code, open_name, raw_cipher, corps_info))

        if abbreviation and abbreviation != open_name:
            abbr_pattern = _build_unit_fuzzy_pattern(abbreviation) if fuzzy_match else re.compile(rf"\b{re.escape(abbreviation)}\b", re.IGNORECASE)
            patterns_to_apply.append((len(abbreviation), abbr_pattern, closed_code, abbreviation, raw_cipher, corps_info))

    # Сортуємо від найдовших до найкоротших
    patterns_to_apply.sort(key=lambda x: x[0], reverse=True)

    for _, pat, fc, op_name, raw_c, c_info in patterns_to_apply:
        matches = pat.findall(body_text)
        if matches:
            replaced_count += len(matches)
            body_text = pat.sub(lambda m, code=fc: _match_case(m.group(0), code), body_text)
            report_rows.append((op_name, raw_c or "(немає)", c_info or "(немає)", fc))

    # 3. Заміна зворотів "цієї самої бригади"
    for pat, replacer in _UNIT_PHRASE_REPLACEMENTS:
        matches = pat.findall(body_text)
        if matches:
            replaced_count += len(matches)
            def _make_phrase_rep(r_func):
                return lambda m: _match_case(m.group(0), r_func(m))
            body_text = pat.sub(_make_phrase_rep(replacer), body_text)

    # 3.1. Нормалізація повторів "військової частини"
    body_text = re.sub(
        r"\b(?:військов(?:ої|а|у|ій|ою)\s+частин(?:и|а|у|і|ою)\s*){2,}",
        "військової частини ",
        body_text,
        flags=re.IGNORECASE,
    )
    body_text = re.sub(
        r"\b(?:ВІЙСЬКОВ(?:ОЇ|А|У|ІЙ|ОЮ)\s+ЧАСТИН(?:И|А|У|І|ОЮ)\s*){2,}",
        "ВІЙСЬКОВОЇ ЧАСТИНИ ",
        body_text,
    )
    body_text = re.sub(
        r"\b(військов\w+\s+частин\w+\s+[АA]\d+)(?:\s+\1\b)+",
        r"\1",
        body_text,
        flags=re.IGNORECASE,
    )

    # 4. Користувацькі правила
    if rules:
        body_text, custom_count = _apply_custom_rules(body_text, rules)
        replaced_count += custom_count

    # 4.1. Повторне згортання дублікатів
    body_text = re.sub(
        r"\b(?:військов(?:ої|а|у|ій|ою)\s+частин(?:и|а|у|і|ою)\s*){2,}",
        "військової частини ",
        body_text,
        flags=re.IGNORECASE,
    )
    body_text = re.sub(
        r"\b(?:ВІЙСЬКОВ(?:ОЇ|А|У|ІЙ|ОЮ)\s+ЧАСТИН(?:И|А|У|І|ОЮ)\s*){2,}",
        "ВІЙСЬКОВОЇ ЧАСТИНИ ",
        body_text,
    )

    # 4.2. Порожні рядки: 1 перед пунктом, 2 перед підписантом
    body_lines_cleaned = []
    for idx, b_line in enumerate(body_text.splitlines()):
        b_clean = b_line.strip()
        is_item = bool(re.match(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+", b_clean))
        is_signer = bool(_ORDER_SIGNER_START_RE.match(b_clean))

        if is_signer and idx > 0 and body_lines_cleaned:
            while body_lines_cleaned and body_lines_cleaned[-1].strip() == "":
                body_lines_cleaned.pop()
            body_lines_cleaned.append("")
            body_lines_cleaned.append("")
        elif is_item and idx > 0 and body_lines_cleaned:
            while len(body_lines_cleaned) > 1 and body_lines_cleaned[-1].strip() == "" and body_lines_cleaned[-2].strip() == "":
                body_lines_cleaned.pop()
            if body_lines_cleaned[-1].strip() != "":
                body_lines_cleaned.append("")
        body_lines_cleaned.append(b_line)
    body_text = "\n".join(body_lines_cleaned)

    # 5. Фінальний текст
    header_str = new_header.strip() if new_header.strip() else ""
    final_text = f"{header_str}\n\n{body_text}".strip() if header_str else body_text.strip()
    table = DataTable(
        ("Відкрита назва", "Закритий код ВЧ", "Армійський корпус", "Форматована заміна у тексті"),
        tuple(report_rows),
        "Звіт частин та корпусів наказу",
    )
    summary = f"Сформовано повідомлення/наказ. Знайдено ВЧ/корпусів: {len(report_rows)}, замін: {replaced_count}"

    return {
        "decision_text": final_text,
        "table": table,
        "replaced_count": replaced_count,
        "summary": summary,
    }
