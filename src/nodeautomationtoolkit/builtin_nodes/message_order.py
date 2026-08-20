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


# ── Відмінювання «військова частина» у зашифрованому змісті ──────────────────
#
# Відкрита назва в наказі стоїть у певному відмінку («до 55 окремого
# батальйону», «у 55 окремому батальйоні»), тому підстановка має ставати
# у той самий відмінок, інакше речення стає неграматичним.
_UNIT_PHRASE_BY_CASE = {
    "Н": "військова частина",
    "Р": "військової частини",
    "Д": "військовій частині",   # давальний і місцевий збігаються
    "З": "військову частину",
    "О": "військовою частиною",
}

# Ознаки відмінка за закінченнями слів у знайденій назві. Перевіряються
# згори вниз; використовуються лише впевнені (переважно прикметникові)
# закінчення. Якщо жодне не спрацювало — лишається родовий відмінок,
# тобто поточна поведінка, і гірше за неї стати не може.
_CASE_ENDING_RULES = (
    (("ого",), "Р"),          # окремого батальйону
    (("ої",), "Р"),           # окремої механізованої бригади
    (("ому",), "Д"),          # окремому батальйону / у окремому батальйоні
    (("ій",), "Д"),           # окремій бригаді
    (("им", "ім"), "О"),      # окремим батальйоном
    (("ою", "ею"), "О"),      # окремою бригадою
    (("ий",), "Н"),           # окремий батальйон
)

# Відмінки, які визначаються лише за узгодженням кількох слів підряд
# (напр. «окрема механізована бригада» — три слова на «-а»).
_CASE_AGREEMENT_RULES = (
    (("а", "я"), "Н"),        # окрема механізована бригада
    (("у", "ю"), "З"),        # окрему механізовану бригаду
)


def _detect_grammatical_case(matched_text: str) -> str:
    """Визначає відмінок знайденої відкритої назви військової частини.

    Повертає одну з міток `_UNIT_PHRASE_BY_CASE`. За відсутності впевнених
    ознак повертає «Р» (родовий) — historично усталену поведінку.
    """
    words = [w for w in re.findall(r"[^\W\d_]+", str(matched_text).lower()) if len(w) > 3]
    if not words:
        return "Р"

    for endings, case_label in _CASE_ENDING_RULES:
        if any(word.endswith(endings) for word in words):
            return case_label

    for endings, case_label in _CASE_AGREEMENT_RULES:
        if sum(1 for word in words if word.endswith(endings)) >= 2:
            return case_label

    return "Р"


def _apply_case_to_closed_text(closed_text: str, case_label: str) -> str:
    """Ставить підстановку «військової частини АXXXX» у потрібний відмінок.

    Змінюється лише ПЕРШЕ входження: якщо частина має корпус, текст має вигляд
    «військової частини АXXXX військової частини АYYYY», де друга половина є
    родовим означенням («частини X частини Y») і має лишатися незмінною.
    Назви ТЦК не містять цього звороту, тому не зачіпаються.
    """
    if not closed_text or case_label == "Р":
        return closed_text
    phrase = _UNIT_PHRASE_BY_CASE.get(case_label)
    if not phrase:
        return closed_text
    return re.sub(
        r"^військової\s+частини", phrase, closed_text, count=1, flags=re.IGNORECASE
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


def find_content_start_line(text: str) -> int:
    """Номер рядка, з якого починається змістовна частина наказу (після шапки)."""
    for idx, line in enumerate(str(text or "").splitlines()):
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
            return idx
    return 0


def cipher_unit_names(
    text: str = "",
    mapping: dict | None = None,
    fuzzy_match: bool = True,
    rules: str | list | dict | None = None,
) -> tuple[str, int, list]:
    """Замінює відкриті назви частин на шифри, НЕ змінюючи структуру рядків.

    На відміну від `generate_decision_order`, не шукає початок змістовної
    частини й не нормалізує порожні рядки. Саме тому придатна для обробки
    ОКРЕМОГО абзацу, скопійованого з наказу разом із форматуванням — так
    зміст повідомлення переноситься тим самим способом, що й у витягах.

    Повертає `(текст, кількість замін, рядки звіту)`.
    """
    if not text:
        return "", 0, []

    text = _fix_military_typos(text)
    mapping_dict = {}
    for key, value in (mapping or {}).items():
        clean_key = _fix_military_typos(str(key))
        if isinstance(value, dict):
            clean_value = dict(value)
            if "open_name" in clean_value:
                clean_value["open_name"] = _fix_military_typos(str(clean_value["open_name"]))
            mapping_dict[clean_key] = clean_value
        else:
            mapping_dict[clean_key] = value

    replaced_count = 0
    report_rows = []

    # 1. Патерни назв частин з урахуванням відмінків
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

        pattern = (
            _build_unit_fuzzy_pattern(open_name)
            if fuzzy_match
            else re.compile(rf"\b{re.escape(open_name)}\b", re.IGNORECASE)
        )
        patterns_to_apply.append((len(open_name), pattern, closed_code, open_name, raw_cipher, corps_info))

        if abbreviation and abbreviation != open_name:
            abbr_pattern = (
                _build_unit_fuzzy_pattern(abbreviation)
                if fuzzy_match
                else re.compile(rf"\b{re.escape(abbreviation)}\b", re.IGNORECASE)
            )
            patterns_to_apply.append(
                (len(abbreviation), abbr_pattern, closed_code, abbreviation, raw_cipher, corps_info)
            )

    # Найдовші назви — першими, щоб коротша не «з'їла» частину довшої.
    patterns_to_apply.sort(key=lambda x: x[0], reverse=True)

    for _, pat, fc, op_name, raw_c, c_info in patterns_to_apply:
        matches = pat.findall(text)
        if matches:
            replaced_count += len(matches)
            text = pat.sub(
                lambda m, code=fc: _match_case(
                    m.group(0),
                    _apply_case_to_closed_text(code, _detect_grammatical_case(m.group(0))),
                ),
                text,
            )
            report_rows.append((op_name, raw_c or "(немає)", c_info or "(немає)", fc))

    # 2. Звороти-посилання: «цієї самої бригади» → «цієї самої військової частини»
    for pat, replacer in _UNIT_PHRASE_REPLACEMENTS:
        matches = pat.findall(text)
        if matches:
            replaced_count += len(matches)

            def _make_phrase_rep(r_func):
                return lambda m: _match_case(m.group(0), r_func(m))

            text = pat.sub(_make_phrase_rep(replacer), text)

    # 3. Згортання повторів «військової частини»
    text = _collapse_unit_phrase_repeats(text)

    # 4. Користувацькі правила
    if rules:
        text, custom_count = _apply_custom_rules(text, rules)
        replaced_count += custom_count
        text = _collapse_unit_phrase_repeats(text)

    return text, replaced_count, report_rows


def _collapse_unit_phrase_repeats(text: str) -> str:
    """Згортає підряд повторені «військової частини» та однакові шифри."""
    text = re.sub(
        r"\b(?:військов(?:ої|а|у|ій|ою)\s+частин(?:и|а|у|і|ою)\s*){2,}",
        "військової частини ",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"\b(?:ВІЙСЬКОВ(?:ОЇ|А|У|ІЙ|ОЮ)\s+ЧАСТИН(?:И|А|У|І|ОЮ)\s*){2,}",
        "ВІЙСЬКОВОЇ ЧАСТИНИ ",
        text,
    )
    return re.sub(
        r"\b(військов\w+\s+частин\w+\s+[АA]\d+)(?:\s+\1\b)+",
        r"\1",
        text,
        flags=re.IGNORECASE,
    )


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
    content_start_idx = find_content_start_line("\n".join(lines))

    body_lines = lines[content_start_idx:]
    body_text = "\n".join(body_lines)

    body_text, replaced_count, report_rows = cipher_unit_names(
        body_text, mapping_dict, fuzzy_match=fuzzy_match, rules=rules
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
