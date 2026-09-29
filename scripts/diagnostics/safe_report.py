"""Безпечний звіт про наказ: СТРУКТУРА без жодного рядка тексту.

Призначення — щоб можна було показати збій розбору стороннім (зокрема мені),
не показуючи сам наказ. У звіт НЕ потрапляє жодного слова з документа: ні ПІБ,
ні назв частин, ні шифрів, ні посад. Виводяться тільки:

* номер рядка, його довжина та формальні ознаки (починається з «N.», закінчується
  двокрапкою, містить «§», частка ВЕЛИКИХ літер, схожість на біографічний рядок);
* як маршрутизація порізала наказ на блоки: діапазони пунктів і шапок;
* адресати — знеособленими мітками Ч-1, Ч-2 … (відповідність зберігається
  ЛОКАЛЬНО в окремий файл, який нікуди не відправляється);
* підсумок: які номери пунктів є в тексті наказу, які дійшли до витягів, а які
  зникли по дорозі.

Запуск:

    python scripts/diagnostics/safe_report.py "наказ.docx" "словник.xlsx"

За замовчуванням звіт друкується в консоль і пишеться поруч у файл
`<наказ>_безпечний_звіт.txt`. Ключ відповідності («Ч-1 = …») пишеться в
`<наказ>_ключ_ЛОКАЛЬНО.txt` — цей файл нікому не надсилайте.
"""
import hashlib
import os
import re
import sys
from importlib import util

sys.stdout.reconfigure(encoding="utf-8")

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(PROJECT_ROOT, "src"))

from nodeautomationtoolkit.builtin_nodes.typography import ORDER_SIGNER_START_RE  # noqa: E402
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (  # noqa: E402
    map_military_units,
    read_recipient_mapping,
)

ITEM_NUMBER_RE = re.compile(r"^\s*(\d{1,3})\s*[.)]")
BIO_RE = re.compile(
    r"(народил|р\.\s*н\.|року народження|у\s+ЗС|ІПН|РНОКПП|ВОС|освіта|підстава)",
    re.IGNORECASE,
)
#: Той самий шаблон, яким генератор шукає підписанта.
SIGNER_RE = ORDER_SIGNER_START_RE
SELF_REF_RE = re.compile(r"\b(?:цього|цієї|того)\s+(?:самого|самої|ж)\b", re.IGNORECASE)


def load_generator():
    """Модуль генератора потрібен лише заради читання тексту наказу тим самим
    способом, яким його читає сам генератор (нумерація рядків має збігатися)."""
    spec = util.spec_from_file_location(
        "generate_extracts_diag", os.path.join(PROJECT_ROOT, "generate_extracts.py")
    )
    module = util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def read_order_text(path: str, generator=None) -> tuple[str, set[int]]:
    """(текст, рядки біографії з відступом від 4 см) — так само, як читає генератор."""
    import win32com.client

    generator = generator or load_generator()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        document = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        try:
            bio_lines: set[int] = set()
            text = generator.read_document_text(document, bio_lines=bio_lines)
            return text, bio_lines
        finally:
            document.Close(False)
    finally:
        word.Quit()


def line_traits(line: str, indented_bio: bool = False) -> str:
    """Формальні ознаки рядка — без жодного його слова."""
    stripped = line.strip()
    if not stripped:
        return "порожній"
    traits = []
    if indented_bio:
        traits.append("ВІДСТУП БІО (≥4 см)")
    if SIGNER_RE.match(stripped):
        traits.append("СХОЖЕ НА ПІДПИСАНТА")
    number = ITEM_NUMBER_RE.match(stripped)
    if number:
        traits.append(f"починається з «{number.group(1)}.»")
    if stripped.startswith("§"):
        traits.append("§")
    if stripped.endswith(":"):
        traits.append("двокрапка в кінці")
    elif stripped.endswith("."):
        traits.append("крапка в кінці")
    letters = [character for character in stripped if character.isalpha()]
    if letters:
        upper_share = sum(character.isupper() for character in letters) / len(letters)
        if upper_share > 0.8:
            traits.append("ВЕЛИКІ літери")
        elif upper_share > 0.3:
            traits.append("частково ВЕЛИКІ")
    if BIO_RE.search(stripped):
        traits.append("схоже на біографічний")
    if SELF_REF_RE.search(stripped):
        traits.append("«цього самого»")
    return ", ".join(traits) or "текст"


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    order_path, mapping_path = sys.argv[1], sys.argv[2]

    generator = load_generator()
    full_text, bio_lines = read_order_text(order_path, generator)
    lines = full_text.splitlines()
    # Так само, як генератор: усе від підписанта відкидається ДО маршрутизації.
    text, _signer = generator.text_before_order_signer(full_text)
    cut_line = len(text.splitlines())
    mapping = read_recipient_mapping(path=mapping_path).get("mapping", {})
    routes = map_military_units(text=text, mapping=mapping, bio_lines=bio_lines)

    # Знеособлення адресатів: стабільна мітка на кожен ключ витягу.
    labels: dict[str, str] = {}

    def label_for(key: str) -> str:
        if key not in labels:
            labels[key] = f"Ч-{len(labels) + 1}"
        return labels[key]

    report: list[str] = []
    digest = hashlib.sha256(os.path.basename(order_path).encode("utf-8")).hexdigest()[:8]
    report.append(f"НАКАЗ: {digest} (хеш назви файлу), рядків: {len(lines)}")
    report.append(
        f"МЕЖА ПІДПИСАНТА: рядок {cut_line} — усе з нього й далі у витяги не йде"
        if cut_line < len(lines) else "МЕЖА ПІДПИСАНТА: не знайдено, текст не обрізано"
    )
    report.append(f"РЯДКІВ БІОГРАФІЇ З ВІДСТУПОМ ≥4 см: {len(bio_lines)}")
    report.append(f"СЛОВНИК: записів {len({id(value) for value in mapping.values()})}")
    report.append("")

    report.append("== СКЛАД ВИТЯГІВ ==")
    numbers_in_extracts: set[int] = set()
    for key, data in routes.get("unit_paragraphs", {}).items():
        items = sorted(
            data.get("items", []), key=lambda value: value.get("source_start_line", 10 ** 9)
        )
        report.append(f"  {label_for(key)}: пунктів {len(items)}")
        for item in items:
            headings = ", ".join(
                f"{heading[0]}–{heading[1]}"
                for heading in (item.get("heading_ranges") or [])
                if isinstance(heading, (list, tuple)) and len(heading) == 2
            )
            number = ITEM_NUMBER_RE.match(str(item.get("label", "")).replace("Пункт", "").strip())
            if number:
                numbers_in_extracts.add(int(number.group(1)))
            report.append(
                f"      {item.get('label', ''):10s} рядки {item.get('source_start_line')}"
                f"–{item.get('source_end_line')}  шапки {headings or '—'}"
            )
    report.append("")

    numbers_in_text = {
        int(match.group(1))
        for match in (ITEM_NUMBER_RE.match(line.strip()) for line in lines)
        if match
    }
    lost = sorted(numbers_in_text - numbers_in_extracts)
    unmatched = {str(item.get("label", "")) for item in routes.get("unmatched_items", [])}
    skipped = {str(item.get("label", "")) for item in routes.get("skipped_items", [])}
    report.append("== ПІДСУМОК ==")
    report.append(f"  номери пунктів у тексті наказу: {sorted(numbers_in_text)}")
    report.append(f"  дійшли до витягів:              {sorted(numbers_in_extracts)}")
    report.append(f"  без адресата (у «Пропущених»):  {sorted(unmatched) or '—'}")
    report.append(f"  виключено (управління):         {sorted(skipped) or '—'}")
    report.append(f"  ЗНИКЛИ БЕЗ СЛІДУ:               {lost or '—'}")
    if lost:
        report.append("  (такий номер є в тексті, але немає ні у витягах, ні в контрольних"
                      " переліках — найчастіше пункт поглинула шапка або сусідній пункт)")
    report.append("")

    if routes.get("unmatched_items"):
        report.append("== ЧОМУ ПУНКТ ЛИШИВСЯ БЕЗ АДРЕСАТА ==")
        from nodeautomationtoolkit.builtin_nodes import recipient_mapping as rm

        for item in routes.get("unmatched_items", []):
            block = str(item.get("text") or "")
            report.append(f"  {item.get('label', '')}: {item.get('reason', '')}")
            report.append(
                f"      ознака ТЦК у тексті: {bool(rm._TCK_KEYWORDS_RE.search(block))}"
                f" | розпізнаний ТЦК: {rm._extract_tck_sender(block)!r}"
            )
            sender = rm._extract_tck_sender(block)
            if sender:
                found = rm._find_entry_in_mapping(sender, sender, mapping)
                report.append(
                    f"      такий рядок у словнику: {'Є' if found else 'НЕМАЄ'}"
                    + (f" (колонка A довж. {len(str(found.get('open_name') or ''))}, "
                       f"B {'заповнена' if str(found.get('cipher') or '').strip() else 'порожня'}, "
                       f"C {'заповнена' if str(found.get('abbreviation') or '').strip() else 'порожня'})"
                       if found else "")
                )
            # Скільки записів словника взагалі зачепилося за текст пункту.
            hits = 0
            for value in {id(v): v for v in mapping.values() if isinstance(v, dict)}.values():
                name = str(value.get("open_name") or "")
                if name and rm._build_unit_fuzzy_pattern(name).search(block):
                    hits += 1
            report.append(f"      записів словника, що збіглися з текстом: {hits}")
            report.append(f"      довжина тексту пункту: {len(block)} символів, рядків: {len(block.splitlines())}")
        report.append("")

    report.append("== РЯДКИ (лише ознаки, без тексту) ==")
    # Усі рядки тіла: пропажа біографії видна саме на рядках МІЖ межами пунктів.
    for index, line in enumerate(lines):
        if index == cut_line:
            report.append("  ---- тут програма обрізала текст (межа підписанта) ----")
        report.append(
            f"  {index:4d} | довж. {len(line):4d} | {line_traits(line, index in bio_lines)}"
        )

    body = "\n".join(report)
    print(body)

    base = os.path.splitext(os.path.abspath(order_path))[0]
    report_path = base + "_безпечний_звіт.txt"
    with open(report_path, "w", encoding="utf-8") as handle:
        handle.write(body + "\n")
    key_path = base + "_ключ_ЛОКАЛЬНО.txt"
    with open(key_path, "w", encoding="utf-8") as handle:
        handle.write("Відповідність міток. НІКОМУ НЕ НАДСИЛАТИ.\n\n")
        for key, label in labels.items():
            handle.write(f"{label} = {key}\n")
    print(f"\nзвіт: {report_path}")
    print(f"ключ (локально, не надсилати): {key_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
