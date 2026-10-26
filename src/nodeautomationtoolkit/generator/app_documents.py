"""Межі тіла наказу та його підписанта у Word.

Винесено з `generate_extracts.py` без змін логіки (26.10.2026).
"""
from __future__ import annotations

import re
from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    _ORDER_BODY_KEYWORDS,
)

from .naming import _slash_to_lines
from .order_text import _DISTRIBUTION_CUTOFF_MARKERS, _find_order_signer
from nodeautomationtoolkit.builtin_nodes.typography import (
    ORDER_SIGNER_START_RE as _ORDER_SIGNER_START_RE,
)


class OrderDocumentMixin:
    """Межі тіла наказу та його підписанта у Word."""

    def _order_body_context(self, source_doc, signer_as_tag: bool = False) -> dict:
        """Межі тіла наказу + реквізити його підписанта для тегів заготовки.

        `signer_as_tag=True` — у заготовці є окремий тег `{{підписант}}`, тож
        блок підписанта підставляється в нього, а `{{зміст}}` завершується
        перед підписантом, щоб не дублювати його.
        """
        parts = self._analyze_order(source_doc)
        body_start = parts["body_start"]
        last_paragraph = parts["last_paragraph"]
        signer_start = parts["signer_start"]
        texts = parts["texts"]

        signature_line = ""
        values: dict[str, str] = {}
        signer = _find_order_signer(source_doc.Content.Text) or {}
        if signer.get("position"):
            values["{{підписант_посада}}"] = _slash_to_lines(signer["position"])
        if signer.get("rank"):
            values["{{підписант_звання}}"] = signer["rank"]
            # Табличний варіант заготовки: звання та прізвище стоять в
            # ОКРЕМИХ комірках, бо вирівняні до різних країв. Пробілами
            # цього не зробити, тому й теги окремі.
            values["{{звання_підписанта}}"] = signer["rank"]
        if signer.get("name"):
            values["{{підписант_піб}}"] = signer["name"]
            values["{{прізвище_підписанта}}"] = signer["name"]

        if signer_as_tag and signer_start and signer_start <= last_paragraph:
            # Блок підписанта беремо з наказу ЯК Є: звання та прізвище в ньому
            # вирівняні пробілами, тому текст переносимо без змін.
            signer_lines = [
                texts[index - 1] for index in range(signer_start, last_paragraph + 1)
            ]
            filled = [line for line in signer_lines if line]
            values["{{підписант}}"] = "\r".join(filled)
            # Останній рядок блоку (звання та прізвище) підкреслюється й
            # притискається праворуч — його треба знайти вже в документі.
            signature_line = filled[-1] if filled else ""

            # Зміст завершується перед підписантом.
            body_end = body_start
            for index in range(signer_start - 1, body_start - 1, -1):
                if texts[index - 1]:
                    body_end = index
                    break
            last_paragraph = body_end

        return {
            "span": (body_start, last_paragraph),
            "values": values,
            "signature_line": signature_line,
            # Чи лишився підписант усередині перенесеного змісту. Від цього
            # залежить нерозривність: інакше останній ПУНКТ вважався б
            # підписантом і втрачав зчеплення з власною шапкою.
            "signer_in_tag": bool(signer_as_tag and signer_start),
        }

    @staticmethod
    def _analyze_order(source_doc) -> dict:
        """Розбирає наказ на межі тіла, підписанта та службової частини.

        Працюємо з АБЗАЦАМИ напряму, а не через зіставлення рядків тексту:
        `Content.Text` не розбиває комірки таблиці на окремі рядки, тоді як
        `doc.Paragraphs` рахує кожну комірку окремо.
        """
        total = source_doc.Paragraphs.Count
        if total < 1:
            raise ValueError("наказ порожній")

        # chr(7) — службовий знак кінця комірки таблиці у Word. Його наявність
        # у тексті абзацу означає, що абзац лежить у таблиці, тож визначити це
        # можна без додаткових звернень до Word.
        raw_texts = [
            (source_doc.Paragraphs(index).Range.Text or "") for index in range(1, total + 1)
        ]
        texts = [raw.replace(chr(7), "").strip() for raw in raw_texts]
        in_table = [chr(7) in raw for raw in raw_texts]

        body_start = 1
        for index, clean in enumerate(texts, start=1):
            if not clean:
                continue
            upper = clean.upper()
            if (
                clean.startswith("§")
                or re.match(r"^\d+[\.\)]", clean)
                or any(keyword in upper for keyword in _ORDER_BODY_KEYWORDS)
            ):
                body_start = index
                break

        def is_service_marker(value: str) -> bool:
            clean = value.casefold()
            return bool(clean) and any(
                clean.startswith(marker) or clean == marker
                for marker in _DISTRIBUTION_CUTOFF_MARKERS
            )

        # Порядок пошуку той самий, що у витягах: спершу з кінця знаходимо
        # службову частину, потім перед нею — підписанта.
        last_marker = total + 1
        for index in range(total, body_start - 1, -1):
            if is_service_marker(texts[index - 1]):
                last_marker = index
                break

        signer_start = None
        for index in range(last_marker - 1, body_start - 1, -1):
            clean = texts[index - 1]
            if clean and _ORDER_SIGNER_START_RE.match(clean):
                signer_start = index
                break

        # Межа копіювання — те, що трапиться раніше ПІСЛЯ підписанта:
        #   • службовий блок («Розрахунок розсилки…» тощо), або
        #   • початок таблиці — це зворот останньої сторінки, його ігноруємо.
        # Таблицю відсікаємо саме за структурою, а не за текстом: у ній може
        # не бути жодного знайомого маркера.
        boundary = total + 1
        for index in range(signer_start or body_start, total + 1):
            if in_table[index - 1] or is_service_marker(texts[index - 1]):
                boundary = index
                break

        # Останній змістовний рядок перед межею — це рядок звання та прізвища
        # підписанта. Абзаци всередині таблиці сюди потрапити не можуть.
        last_paragraph = body_start
        for index in range(boundary - 1, body_start - 1, -1):
            if texts[index - 1] and not in_table[index - 1]:
                last_paragraph = index
                break
        return {
            "body_start": body_start,
            "last_paragraph": last_paragraph,
            "signer_start": signer_start,
            "texts": texts,
        }

    @staticmethod
    def _order_body_span(source_doc) -> tuple[int, int]:
        """Межі тіла наказу разом із підписантом, у номерах абзаців Word."""
        parts = OrderDocumentMixin._analyze_order(source_doc)
        return parts["body_start"], parts["last_paragraph"]
