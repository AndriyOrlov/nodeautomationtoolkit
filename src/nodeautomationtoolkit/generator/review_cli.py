"""Перевірка наказу без вікна (для макросу Word).

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

from . import _paths
import os
import json
from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
)
from nodeautomationtoolkit.builtin_nodes.message_order import (
    find_content_start_line,
)
from nodeautomationtoolkit.order_review import WARNING as REVIEW_WARNING
from nodeautomationtoolkit.order_review import Finding as ReviewFinding
from nodeautomationtoolkit.order_review import review_order
from nodeautomationtoolkit.order_review.check import (
    default_index_folder,
    findings_to_tsv,
)

from .naming import extract_metadata_from_filename
from .order_text import text_before_order_signer
from .units import collect_new_unit_names


# =============================================================================
# ПЕРЕВІРКА НАКАЗУ — спільна для кнопки в програмі й макросу Word
# =============================================================================
def review_order_text(text: str, order_name: str, mapping: dict, index_folder: str):
    """Перевіряє СИРИЙ текст наказу (без `normalize_item_numbering`).

    Номер і дата — з назви файлу, як і скрізь; підписний хвіст відсікається;
    до знахідок `order_review` додаються частини, яких немає в таблиці.
    """
    body, signer = text_before_order_signer(text)
    order_number, order_date = extract_metadata_from_filename(os.path.basename(order_name))
    result = review_order(
        body,
        order_number=order_number,
        order_date=order_date,
        signer=signer,
        mapping=mapping,
        index_folder=index_folder or None,
        full_text=text,  # порожні абзаци перед підписантом — у body їх уже немає
    )
    if mapping:
        content ="\n".join(body.splitlines()[find_content_start_line(body):])
        similar: list = []
        for unit in collect_new_unit_names(content, mapping, similar=similar):
            result.findings.append(ReviewFinding(
                REVIEW_WARNING, "Частини немає в таблиці", "Наказ", f"«{unit}»",
                "Перевірте написання назви або додайте рядок у таблицю.", (unit,),
            ))
        for unit, row in similar:
            result.findings.append(ReviewFinding(
                REVIEW_WARNING, "Частини немає в таблиці", "Наказ",
                f"«{unit}» збігається з рядком таблиці «{row}» не повністю",
                "Перевірте написання в наказі або в стовпці A.", (unit,),
            ))
    return result


def run_review_cli(argv: list[str]) -> int:
    """Консольна перевірка для макросу Word «Перевірити програмою».

    Макрос передає текст відкритого документа (UTF-8, з номерами автонумерації)
    і назву файлу; у відповідь — `findings_to_tsv`. Таблиця й індекс — з config.json
    програми, як у кнопки «🎓 Перевірити наказ». Вікон не показує.
    """
    import argparse

    parser = argparse.ArgumentParser(prog="generate_extracts_qt.py --review-text")
    parser.add_argument("--review-text", required=True)
    parser.add_argument("--order-name", default="")
    parser.add_argument("--out", required=True)
    parser.add_argument("--config", default="")
    parser.add_argument("--index", default=None, help="тека індексу; порожній рядок — без індексу")
    args = parser.parse_args(argv)

    config_path = args.config or os.path.join(_paths.APP_DIR, "config.json")
    config = {}
    try:
        with open(config_path, encoding="utf-8") as handle:
            config = json.load(handle)
    except (OSError, ValueError):
        pass
    notes = []
    mapping = {}
    excel = str(config.get("excel_path") or "")
    if excel and os.path.isfile(excel):
        mapping = read_recipient_mapping(path=excel).get("mapping", {})
    else:
        notes.append("Таблицю відповідностей не вибрано в програмі — адресати пунктів не перевірялись.")
    if args.index is not None:
        index_folder = args.index
    else:
        index_folder = default_index_folder(str(config.get("new_order_index_folder") or ""))

    with open(args.review_text, encoding="utf-8-sig") as handle:
        text = handle.read()
    result = review_order_text(text, args.order_name or "наказ.docx", mapping, index_folder)
    result.notes[:0] = notes
    with open(args.out, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(findings_to_tsv(result))
    return 0
