"""Текст наказу: підписант, службовий хвіст, компонування «2 на 1».

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import re
from nodeautomationtoolkit.builtin_nodes.typography import (
    ORDER_SIGNER_START_RE as _ORDER_SIGNER_START_RE,
)



#: Перелік звань окремо: він потрібен і прив'язаним до початку рядка (звичайний
#: підписний блок), і всередині рядка (коли звання стоїть поруч із посадою).
_RANK_ALTERNATIVES = (
    r"генерал(?:[-\s](?:майор|лейтенант|полковник))?|адмірал(?:[-\s]\w+)?|"
    r"полковник|підполковник|майор|капітан(?:[-\s](?:лейтенант|[1-3]\s+рангу))?|"
    r"старший\s+лейтенант|молодший\s+лейтенант|лейтенант|головний\s+сержант|"
    r"штаб[-\s]сержант|майстер[-\s]сержант|старший\s+сержант|молодший\s+сержант|"
    r"сержант|старшина|солдат|матрос"
)


_ORDER_SIGNER_RANK_RE = re.compile(
    r"^\s*(" + _RANK_ALTERNATIVES + r")\b",
    re.IGNORECASE | re.UNICODE,
)


_ORDER_SIGNER_RANK_INLINE_RE = re.compile(
    r"\b(" + _RANK_ALTERNATIVES + r")\b",
    re.IGNORECASE | re.UNICODE,
)


# Пронумерований пункт наказу: «1.», «2.3.», «10)». Та сама форма, що вже
# використовується в recipient_mapping.py.
_ITEM_START_RE = re.compile(r"^\d{1,3}(?:\.\d{1,3})*[\.\)]\s+")


def _last_item_line(lines: list[str]) -> int:
    """Номер рядка ОСТАННЬОГО пронумерованого пункту наказу.

    Підписант завжди йде після пунктів, тому цей рядок — природна нижня межа
    пошуку підписанта. Вона не залежить від довжини службового хвоста, на
    відміну від лічильника рядків, який доводилося збільшувати щоразу, коли
    в наказі траплялася більша таблиця розсилки.

    Подавати сюди треба ЛИШЕ тіло наказу, до маркера розсилки: рядки самої
    таблиці розсилки теж бувають пронумеровані («1. в/ч А0000 — 1 прим.»), і
    тоді межа заїхала б за підписанта, а його пошук не дав би нічого.
    """
    for index in range(len(lines) - 1, -1, -1):
        if _ITEM_START_RE.match(lines[index]):
            return index
    return 0


def order_extracts_for_2up(extracts: list) -> list:
    """Порядок витягів для друку «2 на 1», який економить папір.

    Порядок витягів значення не має, а зайва порожня сторінка — це зайвий
    папір (користувач, 26.10.2026). Багатосторінковий витяг завжди починається
    з нового аркуша, а після непарного — порожня сторінка (зворот його
    останнього аркуша чужим витягом не займається). Порожня сторінка ПЕРЕД
    багатосторінковим виникала лише тоді, коли перед ним стояла непарна
    кількість одинарних. Тому спершу всі багатосторінкові, потім одинарні:
    кожен багатосторінковий починається з нового аркуша сам, а одинарні йдуть
    парами без пропусків. Усередині груп зберігається вихідний порядок.

    `extracts` — кортежі, де другий елемент — кількість сторінок.
    """
    multis = [item for item in extracts if item[1] > 1]
    singles = [item for item in extracts if item[1] <= 1]
    return multis + singles


def plan_2up_page_layout(extract_pages: list[int]) -> list[dict]:
    """
    Розраховує розміщення витягів та вставку порожніх сторінок для друку «2 сторінки на 1 аркуш».
    Кожен фізичний аркуш містить 2 логічні сторінки (ліва - непарна, права - парна).
    Багатосторінкові витяги завжди починаються з нового аркуша і, якщо їхня довжина непарна,
    доповнюються порожньою сторінкою для вирівнювання наступного аркуша.
    """
    if not extract_pages:
        return []

    events = []
    current_doc_pages = 0

    for i, p_count in enumerate(extract_pages):
        # Якщо витяг багатосторінковий і позиція на непарній сторінці -> вставляємо порожню перед
        if p_count > 1 and (current_doc_pages % 2 != 0):
            events.append({"action": "insert_blank_before", "extract_idx": i})
            current_doc_pages += 1

        events.append({
            "action": "insert_extract",
            "extract_idx": i,
            "pages": p_count,
            "start_page": current_doc_pages + 1,
            "end_page": current_doc_pages + p_count,
        })
        current_doc_pages += p_count

        # Якщо витяг багатосторінковий і непарної довжини -> вставляємо порожню після
        if p_count > 1 and (p_count % 2 != 0):
            events.append({"action": "insert_blank_after", "extract_idx": i})
            current_doc_pages += 1

    return events


_DISTRIBUTION_CUTOFF_MARKERS = (
    "розрахунок розсилки",
    "таблиця розсилки",
    "список розсилки",
    "розсилка:",
    "відмітки служби діловодства",
    "служба діловодства",
    "згідно з оригіналом",
    "розіслано:",
    "відмітка про виконання",
)


def find_distribution_cutoff_line(text: str) -> int:
    """Номер рядка, з якого починається службова таблиця розсилки/відміток.

    Усе до цього рядка — власне наказ разом із підписантом. Пошук іде з кінця
    документа, щоб випадкове входження маркера в тілі пункту не обрізало наказ
    передчасно.
    """
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]

    # Документ сканується ПОВНІСТЮ: цикл іде з кінця й повертає ОСТАННІЙ
    # маркер, тож обмежувати глибину не потрібно. Раніше тут стояло вікно на
    # кілька сотень рядків, і воно ламалося щоразу, коли службовий хвіст
    # виростав: відколи текст збирається з абзаців (`read_document_text`),
    # кожна комірка таблиці розсилки стала окремим рядком. Маркер лишався поза
    # вікном, підписант визначався неправильно, а останній пункт «затягував»
    # службовий хвіст у витяг.
    for idx in range(len(lines) - 1, -1, -1):
        line_lower = lines[idx].casefold().strip()
        if any(line_lower.startswith(marker) or marker == line_lower for marker in _DISTRIBUTION_CUTOFF_MARKERS):
            return idx
    return len(lines)


def find_service_block_line(text: str, start_line: int = 0) -> int:
    """Номер рядка ПЕРШОГО службового блоку, починаючи з `start_line`.

    На відміну від `find_distribution_cutoff_line`, яка сканує з кінця й
    повертає ОСТАННІЙ маркер, ця функція йде вперед. Для примірника потрібен
    саме перший блок після підписанта: інакше проміжні таблиці на кшталт
    «Розрахунок розсилки витягів із наказу» лишалися всередині документа.
    """
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip().casefold() for line in raw_lines]
    for index in range(max(0, start_line), len(lines)):
        clean = lines[index]
        if not clean:
            continue
        if any(clean.startswith(marker) or clean == marker for marker in _DISTRIBUTION_CUTOFF_MARKERS):
            return index
    return len(lines)


def _find_order_signer(text: str) -> dict[str, str] | None:
    """Повертає реквізити й номер рядка початку підписанта (Командувача) в наказі, відсікаючи таблицю розсилки."""
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]

    reference_index = find_distribution_cutoff_line(text)
    # Пошук іде з кінця й зупиняється на останньому пронумерованому пункті:
    # підписант стоїть після пунктів, а не всередині них. Це водночас знімає
    # залежність від довжини службового хвоста і не дає прийняти за підписанта
    # рядок тіла наказу, що починається з «Начальник…»/«Командир…».
    #
    # Межу пунктів рахуємо до ПЕРШОГО службового блоку, а не до останнього
    # маркера: у зразку звороту останнього аркуша (додаток 43) службових
    # блоків два — «Розрахунок розсилки витягів із наказу:» і «Розрахунок
    # розсилки електронних повідомлень:», — а їхні рядки теж пронумеровані
    # («1. Військова частина А0000 п. 1.»). Рахуючи їх пунктами наказу, межа
    # заїжджала за підписанта, той не знаходився зовсім, і в витяг протікав
    # увесь службовий хвіст.
    body_limit = min(find_service_block_line(text, 0), max(0, reference_index))
    search_start = _last_item_line(lines[:body_limit])

    # Після останнього пункту може бути кілька підписоподібних службових
    # блоків. Для межі тіла наказу потрібен ПЕРШИЙ справжній підписант.
    # Зворотний пошук вибирав останній блок і затягував усе між ними в текст
    # останнього пункту витягу.
    for start_index in range(search_start + 1, reference_index):
        if not _ORDER_SIGNER_START_RE.match(lines[start_index]):
            continue

        position_lines = []
        # Межа блоку — початок службової частини, а не лічильник рядків.
        # Вікно у 8 рядків рвалося, щойно між посадою та званням
        # ставало більше восьми абзаців — а порожніми абзацами підписний
        # блок у наказі часто розсувають до низу сторінки. Підписант
        # тоді не знаходився ЗОВСІМ, тіло наказу не обрізалося, і блок
        # протікав у {{зміст}}.
        for line_index in range(start_index, reference_index):
            line = lines[line_index]
            if not line:
                continue
            rank_match = _ORDER_SIGNER_RANK_RE.match(line)
            if rank_match:
                rank = rank_match.group(1)
                remainder = line[rank_match.end():].strip(" .\t–—-:")
                columns = [
                    part.strip(" .\t–—-:")
                    for part in re.split(r"\t+|\.{3,}", remainder)
                    if part.strip(" .\t–—-:")
                ]
                name = columns[-1] if columns else remainder
                if not name:
                    for next_index in range(line_index + 1, min(reference_index, line_index + 5)):
                        candidate = lines[next_index]
                        if any(marker in candidate.casefold() for marker in _DISTRIBUTION_CUTOFF_MARKERS) or _ORDER_SIGNER_START_RE.match(candidate):
                            break
                        candidate_words = re.findall(r"[A-Za-zА-Яа-яІіЇїЄєҐґ'’`-]+", candidate)
                        if len(candidate_words) >= 2 and candidate.casefold() not in {"підпис", "м. п."}:
                            name = " ".join(candidate_words)
                            break
                return {
                    "start_line": start_index,
                    "position": "\n".join(position_lines),
                    "rank": rank,
                    "name": name,
                }
            if any(marker in line.casefold() for marker in _DISTRIBUTION_CUTOFF_MARKERS):
                break
            if any(char.isalpha() for char in line) and line.casefold() not in {"підпис", "м. п."}:
                position_lines.append(line)

        # Фолбек: звання стоїть не з початку рядка, а в тому ж рядку, що й
        # посада («Командувач військ … генерал-лейтенант Іван ПРІЗВИЩЕНКО»).
        # Без цього блок не впізнавався ЗОВСІМ: тіло наказу не обрізалося, а
        # підписант не переносився у витяг.
        for line_index in range(start_index, reference_index):
            line = lines[line_index]
            if not line:
                continue
            if any(marker in line.casefold() for marker in _DISTRIBUTION_CUTOFF_MARKERS):
                break
            inline_rank = _ORDER_SIGNER_RANK_INLINE_RE.search(line)
            if not inline_rank or inline_rank.start() == 0:
                continue
            tail = line[inline_rank.end():].strip(" .\t–—-:")
            name_words = re.findall(r"[A-Za-zА-Яа-яІіЇїЄєҐґ'’`-]+", tail)
            if len(name_words) < 2:
                continue
            return {
                "start_line": start_index,
                "position": line[: inline_rank.start()].strip(" .\t–—-:"),
                "rank": inline_rank.group(1),
                "name": " ".join(name_words),
            }
    return None


def extract_order_signer(text: str) -> dict[str, str]:
    """Виділяє посаду, звання й ПІБ багаторядкового підписанта наказу."""
    signer = _find_order_signer(text)
    if not signer:
        return {"position": "", "rank": "", "name": ""}
    return {key: signer.get(key, "") for key in ("position", "rank", "name")}


def _first_signer_like_line_after_last_item(text: str) -> int | None:
    """Перша підписоподібна межа після останнього пункту.

    Тимчасове жорстке правило: такий блок ніколи не є продовженням останнього
    пункту, навіть якщо в ньому не вдалося розібрати звання та ПІБ.
    """
    raw_lines = str(text or "").replace("\x07", "").splitlines()
    lines = [re.sub(r"\s+", " ", line).strip() for line in raw_lines]
    body_limit = find_service_block_line(text, 0)
    last_item = _last_item_line(lines[:body_limit])
    for index in range(last_item + 1, body_limit):
        if _ORDER_SIGNER_START_RE.match(lines[index]):
            return index
    return None


def text_before_order_signer(text: str) -> tuple[str, dict[str, str]]:
    """Відсікає підписанта й увесь службовий текст, який іде після нього."""
    signer = _find_order_signer(text)
    signer_like_start = _first_signer_like_line_after_last_item(text)
    if signer_like_start is None and not signer:
        return text, {"position": "", "rank": "", "name": ""}
    raw_lines = str(text or "").splitlines()
    start_line = signer_like_start if signer_like_start is not None else signer["start_line"]
    # Не підтягуємо реквізити з пізнішого підписоподібного блока. Вони
    # прийнятні лише тоді, коли належать саме першій відсіченій межі.
    clean_signer = (
        {key: signer.get(key, "") for key in ("position", "rank", "name")}
        if signer and signer.get("start_line") == start_line
        else {"position": "", "rank": "", "name": ""}
    )
    return "\n".join(raw_lines[:start_line]).rstrip(), clean_signer
