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
    is_tck_entry,
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


# ── Запобіжник від проковтування тексту ─────────────────────────────────────
#
# У пункті наказу «звідки» пишеться малими, а «КУДИ» — ВЕЛИКИМИ. Назва однієї
# частини НІКОЛИ не буває наполовину малою, наполовину ВЕЛИКОЮ. Якщо збіг
# містить і те, і те — він перетнув межу пункту й з'їв текст між ними, тому
# такий збіг відхиляється й текст лишається як був.
#
# Втрата тексту наказу — найгірший можливий наслідок, тому запобіжників два:
# цей і заборона переходити тире-роздільник у самому патерні
# (`_build_unit_fuzzy_pattern`, розд. 4.2.9).
_MATCH_WORD_RE = re.compile(r"[^\W\d_]+", re.UNICODE)


def _spans_source_and_destination(matched_text: str) -> bool:
    """Чи перетнув збіг межу «звідки → КУДИ» (мішанина регістрів)."""
    words = _MATCH_WORD_RE.findall(str(matched_text or ""))
    # Довжина відсікає абревіатури (ТЦК, СП, АК, НГУ) — вони бувають ВЕЛИКИМИ
    # всередині цілком правильної назви й межі пункту не позначають.
    has_upper_word = any(len(word) >= 5 and word.isupper() for word in words)
    has_lower_word = any(len(word) >= 4 and word.islower() for word in words)
    return has_upper_word and has_lower_word


# ── Почесне найменування в лапках після зашифрованої частини ────────────────
#
# Почесне найменування («Едельвейс», «Холодний Яр») є частиною ВІДКРИТОЇ назви
# частини й однозначно її ідентифікує. Патерн словника закінчується на роді
# частини («…гірсько-штурмової бригади»), тому лапки лишалися в тексті вже
# ПІСЛЯ шифру: «військової частини А0000 “Едельвейс”». Це і збій вимоги
# (зайвий текст), і витік відкритої ознаки в закритому повідомленні.
_CLOSED_UNIT_PHRASE = r"військов\w+\s+частин\w+\s+[АA]\s?\d+"
_HONORIFIC_AFTER_CLOSED_UNIT_RE = re.compile(
    rf"((?:{_CLOSED_UNIT_PHRASE})(?:\s+{_CLOSED_UNIT_PHRASE})*)"
    r"\s*[«“„\"]([^«»“”„\"]{1,80})[»”\"]",
    re.IGNORECASE | re.UNICODE,
)


def _strip_honorific_after_closed_unit(text: str) -> str:
    """Прибирає почесне найменування в лапках одразу після шифру частини."""
    previous = None
    result = str(text or "")
    # Кілька проходів: «…А0000 “Х” “У”» трапляється у переліках.
    while previous != result:
        previous = result
        result = _HONORIFIC_AFTER_CLOSED_UNIT_RE.sub(r"\1", result)
    return result


# ── Мʼякі переноси (Shift+Enter) у змісті повідомлення ──────────────────────
#
# У наказі довгу відкриту назву часто рве Shift+Enter. У повідомленні на її
# місці лишається три слова («військової частини АXXXX»), і рядок обривається
# посеред фрази — половина рядка порожня. Тому в ПОВІДОМЛЕННЯХ (і тільки в
# них — витяг і примірник копіюють наказ 1-в-1) перенос зшивається.
#
# Правило закінчення: перенос лишається там, де рядок ЗАВЕРШЕНО — після коми,
# крапки, двокрапки тощо. Саме так побудований біографічний блок
# («…освіта: ТВІ у 2012 р.,» ⏎ «у ЗС - із 08.2008.»), і його структура не
# змінюється. Розрив посеред фрази прибирається, і абзац переливається сам.
_SOFT_BREAK_RE = re.compile("[ \t]*[\x0b\n][ \t]*")
_KEEP_BREAK_AFTER = ",.;:!?…»”\")"


def reflow_soft_breaks(text: str) -> str:
    """Зшиває мʼякі переноси посеред фрази, лишаючи переноси після закінчення."""
    source = str(text or "")

    def _replace(match: re.Match) -> str:
        head = source[: match.start()].rstrip()
        if head and head[-1] in _KEEP_BREAK_AFTER:
            return match.group(0)
        return " "

    return _SOFT_BREAK_RE.sub(_replace, source)


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
        if is_tck_entry(mapped_val) or is_tck_entry(open_name):
            # ТЦК не шифрується: у змісті його назва лишається ПОВНОЮ
            # відкритою (розд. 9.5.6). Підстановка короткої форми зі словника
            # («Тестовий ОТЦК та СП») сенсу не мала, а рядки ТЦК — найдовші
            # в словнику, тож саме вони й проковтували текст пункту (4.2.9).
            # Для маршрутизації ці рядки далі потрібні — там вони не змінені.
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
        hits = [0]

        def _replace(match, code=fc, hits=hits):
            matched = match.group(0)
            if _spans_source_and_destination(matched):
                return matched  # збіг перетнув межу пункту — не чіпаємо текст
            hits[0] += 1
            return _match_case(
                matched,
                _apply_case_to_closed_text(code, _detect_grammatical_case(matched)),
            )

        text = pat.sub(_replace, text)
        if hits[0]:
            replaced_count += hits[0]
            report_rows.append((op_name, raw_c or "(немає)", c_info or "(немає)", fc))

    # 1.1. Почесне найменування в лапках лишається після шифру — прибираємо
    # його ДО згортання повторів, інакше воно розділяє два однакові шифри
    # («…А0000 “Едельвейс” військової частини А0000») і повтор не згортається.
    text = _strip_honorific_after_closed_unit(text)

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
    """Згортає підряд повторені «військової частини» та однакові шифри.

    Лишається ПЕРШИЙ зворот — саме він несе відмінок речення (його поставив
    `_apply_case_to_closed_text`) і регістр («ВІЙСЬКОВОЇ ЧАСТИНИ» у частині,
    КУДИ призначають, лишається великим). Раніше повтор завжди згортався в
    малий родовий, тож «у військовій частині військовій частині А0000»
    ставало неграматичним «військової частини А0000».
    """
    unit_phrase = r"військов(?:ої|а|у|ій|ою)\s+частин(?:и|а|у|і|ою)"
    text = re.sub(
        rf"\b({unit_phrase})(?:\s+{unit_phrase})+\s*",
        r"\1 ",
        text,
        flags=re.IGNORECASE,
    )
    return _collapse_unit_chain_repeats(text)


# Ланка ланцюга підпорядкованості: «військової частини АXXXX».
_CHAIN_LINK_RE = re.compile(
    r"(військов\w+\s+частин\w+)\s+([АA]\s?\d+)", re.IGNORECASE | re.UNICODE
)
_CHAIN_RUN_RE = re.compile(
    rf"{_CLOSED_UNIT_PHRASE}(?:\s+{_CLOSED_UNIT_PHRASE})+", re.IGNORECASE | re.UNICODE
)


def _collapse_unit_chain_repeats(text: str) -> str:
    """Прибирає повтор шифру в ланцюгу підпорядкованості.

    Ланцюг іде від меншого до більшого: батальйон → бригада → корпус, і
    **три ланки — це нормально**, якщо батальйон має власний номер і шифр.
    Прибирається лише ПОВТОР того самого шифру, і лишається його **остання**
    поява: більше зʼєднання завжди стоїть далі.

    Звідки береться повтор: стовпець D так і називається «Корпус», тому в
    рядку батальйону там часто стоїть корпус, а не бригада. Тоді і батальйон,
    і бригада тягнуть за собою ту саму ланку корпусу, і виходило
    «А1111 А3333 А2222 А3333» замість «А1111 А2222 А3333».

    Відмінок і регістр речення несе ПЕРША ланка (`_apply_case_to_closed_text`),
    тому її написання переноситься на першу ланку, що лишилась.
    """

    def _collapse_run(match: re.Match) -> str:
        links = _CHAIN_LINK_RE.findall(match.group(0))
        if len(links) < 2:
            return match.group(0)

        last_position = {}
        for index, (_phrase, cipher) in enumerate(links):
            last_position[re.sub(r"\s+", "", cipher).upper()] = index
        kept = [links[index] for index in sorted(set(last_position.values()))]
        if len(kept) == len(links):
            # Повторів немає — віддаємо збіг ЯК Є. Інакше складання через
            # пробіл затерло б мʼякий перенос (Shift+Enter), яким у наказі
            # часто розірвано назву частини (правило 4.2.8).
            return match.group(0)

        # Написання першої ланки (відмінок, ВЕЛИКІ літери) належить реченню,
        # а не конкретному шифру, тому лишається на першому місці.
        kept[0] = (links[0][0], kept[0][1])
        return " ".join(f"{phrase} {cipher}" for phrase, cipher in kept)

    return _CHAIN_RUN_RE.sub(_collapse_run, text)


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
