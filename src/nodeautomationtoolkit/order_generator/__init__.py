"""Генератор наказів по особовому складу (вкладка «0. Накази»).

Чотири етапи, кожен окремим модулем (`pipeline.py` їх зшиває):

1. **Індексація** (`..order_index`) — папка зі старими наказами → локальний
   індекс посад, підрозділів і пунктів про осіб.
2. **Генерація** — джерела в будь-якому наборі (план переміщення `plan.py`,
   документ чи скан `documents.py` + `ocr.py`, пункт попереднього наказу з
   індексу, ручний ввід) → запис про особу (`record.py`, `sources.py`,
   `merge.py`, `resolve.py`) → текст наказу за зразками додатка 53
   (`compose.py`, ПІБ у знахідному — `names.py`).
3. **Перевірка** (`check.py`) — що бракує, що не сходиться, де відмінок під
   сумнівом; помилки не дають збирати документ.
4. **Збірка** (`build.py`) — .docx за правилами верстки витягів
   (PROJECT_RULES розд. 5: шрифти, відступи, порожні рядки, нерозривність).

Усе локально (PROJECT_RULES 1.1). Шаблон наказу («Шаблон наказу» в
налаштуваннях вкладки) необов'язковий: без нього аркуш будується з нуля.
"""

from .build import OrderDocumentParts, build_order_document, order_filename
from .check import CheckResult, Problem, check_draft
from .compose import OrderDraft, OrderItem, OrderParams, compose_order
from .record import PersonRecord

__all__ = [
    "CheckResult",
    "OrderDocumentParts",
    "OrderDraft",
    "OrderItem",
    "OrderParams",
    "PersonRecord",
    "Problem",
    "build_order_document",
    "check_draft",
    "compose_order",
    "order_filename",
]
