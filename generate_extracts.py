# Інтерфейс будує Qt-оболонка (`generate_extracts_qt.py`, `generator_qt`). Логіка
# написана під API tkinter (StringVar, messagebox, filedialog, константи), і оболонка
# перед запуском підміняє ці імена своєю реалізацією (`install_qt_bridge`).
from __future__ import annotations

# Фасад: імпорти тут — реекспорт старих імен для тестів і скриптів.
# ruff: noqa: F401, F403

import os
import sys
import hashlib
import json
import re
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.constants import *

import win32com.client
import openpyxl

# Додаємо src до sys.path для імпорту модулів nodeautomationtoolkit
project_root = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(project_root, "src")
if src_path not in sys.path:
    sys.path.insert(0, src_path)

# Номер версії — один на всю програму (`src/nodeautomationtoolkit/__init__.py`).
# Домовленість із користувачем 18.09.2026: КОЖНА наша зміна в програмі збільшує
# останнє число на одиницю, щоб у вікні було видно, чи запущена вже виправлена
# версія, чи ще стара.
from nodeautomationtoolkit import __version__ as APP_VERSION

from nodeautomationtoolkit.builtin_nodes.recipient_mapping import (
    read_recipient_mapping,
    _build_unit_fuzzy_pattern,
    keep_open_names_of,
    _table_cipher,
    map_military_units,
    normalize_item_numbering,
    _format_full_closed_unit_text,
    _format_item_numbers_range,
)
from nodeautomationtoolkit.builtin_nodes.message_order import (
    generate_decision_order,
    cipher_unit_names,
    clear_table_caches,
    find_content_start_line,
    reflow_soft_breaks,
)
from nodeautomationtoolkit.builtin_nodes.copy_generator import (
    PREVIEW_DEFAULT_DELAY,
    PreviewSteps,
    build_copy_document,
    retry_on_busy_word,
    _ORDER_BODY_KEYWORDS,
)
from nodeautomationtoolkit.builtin_nodes.blank_images import remove_blank_images
from nodeautomationtoolkit.order_review import WARNING as REVIEW_WARNING
from nodeautomationtoolkit.order_review import Finding as ReviewFinding
from nodeautomationtoolkit.order_review import review_order
from nodeautomationtoolkit.order_review.check import (
    default_index_folder,
    findings_to_tsv,
    group_findings,
)
from nodeautomationtoolkit.order_review.marking import save_marked_copy
from nodeautomationtoolkit.builtin_nodes.template_tags import (
    SIGNER_TAGS, tag_aliases, expand_common_tags, certifier_tags, signer_tags,
)



#: Правила набору — спільні з генератором наказів.
from nodeautomationtoolkit.builtin_nodes.typography import (  # noqa: F401
    ORDER_SIGNER_START_RE as _ORDER_SIGNER_START_RE,
    apply_ukrainian_typography,
    clean_duplicated_units,
    ensure_blank_line_before_items,
)

# ── Логіка винесена в пакет `nodeautomationtoolkit.generator` (26.10.2026) ──
# Цей файл лишився точкою входу й фасадом: усі старі імена доступні тут, як і
# раніше (`generate_extracts.read_document_text`, `generate_extracts.App`…).
from nodeautomationtoolkit.generator import _paths as _generator_paths

# `config.json` — поруч із цим файлом, як і до розбиття (і в EXE теж).
_generator_paths.APP_DIR = os.path.dirname(os.path.abspath(__file__))

from nodeautomationtoolkit.generator.naming import (  # noqa: F401
    COPY_FILENAME_PREFIX,
    _LEGACY_COPY_PREFIX,
    _slash_to_lines,
    back_page_tag_values,
    build_copy_two_filename,
    build_extracts_filename,
    extract_metadata_from_filename,
    extract_metadata_from_text,
    format_message_date,
    format_ukr_date,
    is_generated_copy_filename,
    sanitize_filename,
)
from nodeautomationtoolkit.generator.errors import (  # noqa: F401
    UserError,
    _LOG_MEASURE_WORDS,
    _LOG_NAME_WORD,
    _LOG_SURNAME_UPPER,
    _LOG_UPPER_KEEP,
    _REDACT_RE,
    _WINDOWS_BUSY_ERRORS,
    _WINDOWS_DRIVE_ERRORS,
    _error_file_name,
    explain_error,
    redact_sensitive_text,
)
from nodeautomationtoolkit.generator.word_doc import (  # noqa: F401
    BIO_MIN_LEFT_INDENT_PT,
    DEATH_ITEM_FONT_SIZES,
    EXTRA_BLANK_SPACINGS,
    EXTRA_TEXT_SPACINGS,
    WD_UNDEFINED,
    _BATCH_WORD_SETTINGS,
    _DEATH_ITEM_RE,
    _LIST_ITEM_NUMBER_RE,
    _apply_carried_formatting,
    _carry_source_formatting,
    _list_item_numbers_by_start,
    _read_source_formatting,
    _word_setting_target,
    copy_template_for_editing,
    detect_word_extension,
    force_quit_word,
    is_biographical_paragraph,
    is_death_item_text,
    is_page_break_only_text,
    is_path_writable,
    iter_paragraphs,
    last_paragraph,
    paragraph_line_count,
    read_document_text,
    remove_page_breaks_in_range,
    restore_word_settings,
    tune_word_for_batch,
)
from nodeautomationtoolkit.generator.reports import (  # noqa: F401
    CONTROL_REPORTS_MAX_AGE_DAYS,
    MANAGEMENT_RESULT_LABEL,
    _control_reports_cleaned,
    _save_table_to_excel,
    _saved_xlsx_rows,
    cleanup_control_reports,
    control_report_path,
    control_reports_folder,
    load_saved_analysis_reports,
    management_calculation_rows,
    select_extracts_for_scope,
)
from nodeautomationtoolkit.generator.units import (  # noqa: F401
    _NUMBERED_UNIT_GAP_WORD,
    _NUMBERED_UNIT_KIND,
    _QUOTED_NAME_AFTER_RE,
    _UNIT_ALREADY_CLOSED_AFTER_RE,
    _UNIT_KIND_KEYS,
    _UNIT_NAME_TAIL_WORD_RE,
    _UNMATCHED_OPEN_UNIT_RE,
    _append_rows_with_excel,
    _quoted_name_length,
    _similar_table_row,
    _table_open_name_column,
    _unit_name_tail_length,
    _unit_number_and_kind,
    append_unit_stubs_to_table,
    collect_new_unit_names,
    describe_cipher_problems,
    find_unmatched_open_unit_spans,
    unmatched_open_unit_spans,
)
from nodeautomationtoolkit.generator.messages import (  # noqa: F401
    MESSAGE_BY_CALCULATION_TEXT,
    _MESSAGE_ADDRESSEE_KIND_TAGS,
    _MESSAGE_ITEM_BREAK_RE,
    build_addressee_kind_text,
    build_message_recipient_groups,
    build_message_recipient_list,
    message_item_contexts,
    message_skipped_item_lines,
)
from nodeautomationtoolkit.generator.order_text import (  # noqa: F401
    _DISTRIBUTION_CUTOFF_MARKERS,
    _ITEM_START_RE,
    _ORDER_SIGNER_RANK_INLINE_RE,
    _ORDER_SIGNER_RANK_RE,
    _RANK_ALTERNATIVES,
    _find_order_signer,
    _first_signer_like_line_after_last_item,
    _last_item_line,
    extract_order_signer,
    find_distribution_cutoff_line,
    find_service_block_line,
    order_extracts_for_2up,
    plan_2up_page_layout,
    text_before_order_signer,
)
from nodeautomationtoolkit.generator.review_cli import (  # noqa: F401
    review_order_text,
    run_review_cli,
)
from nodeautomationtoolkit.generator.app_core import (  # noqa: F401
    CoreMixin,
)
from nodeautomationtoolkit.generator.app_documents import (  # noqa: F401
    OrderDocumentMixin,
)
from nodeautomationtoolkit.generator.app_orders import (  # noqa: F401
    OrdersMixin,
)
from nodeautomationtoolkit.generator.app_extracts import (  # noqa: F401
    ExtractsMixin,
)
from nodeautomationtoolkit.generator.app_review import (  # noqa: F401
    ReviewMixin,
)
from nodeautomationtoolkit.generator.app_messages import (  # noqa: F401
    MessagesMixin,
)
from nodeautomationtoolkit.generator.app_printing import PrintingMixin  # noqa: F401
from nodeautomationtoolkit.generator.app_copies import (  # noqa: F401
    CopiesMixin,
)


class App(
    CoreMixin,
    OrderDocumentMixin,
    OrdersMixin,
    ExtractsMixin,
    ReviewMixin,
    MessagesMixin,
    CopiesMixin,
    PrintingMixin,
):
    """Генератор витягів, примірників і повідомлень — збирається з міксинів пакета."""


# ── Підміни передаються в пакет ─────────────────────────────────────────────
# Qt-оболонка підміняє тут `tk`, `messagebox`, `filedialog`, `DISABLED`, `NORMAL`
# (`install_qt_bridge`), тести — окремі функції (`force_quit_word`…). Код тепер
# живе в модулях пакета й бере ці імена зі СВОЇХ глобальних, тому присвоєння
# атрибута цьому модулю дублюється в кожен модуль пакета, де таке ім'я є.
import gc as _gc  # noqa: E402
import types as _types  # noqa: E402

from nodeautomationtoolkit.generator import (  # noqa: E402
    naming,
    errors,
    word_doc,
    reports,
    units,
    messages,
    order_text,
    review_cli,
    app_core,
    app_documents,
    app_orders,
    app_extracts,
    app_review,
    app_messages,
    app_copies,
    app_printing,
    printing,
)

_GENERATOR_PARTS = (naming, errors, word_doc, reports, units, messages, order_text, review_cli, app_core, app_documents, app_orders, app_extracts, app_review, app_messages, app_copies, app_printing, printing,)


class _FacadeModule(_types.ModuleType):
    def __setattr__(self, name, value):
        super().__setattr__(name, value)
        for part in _GENERATOR_PARTS:
            if name in part.__dict__:
                setattr(part, name, value)


def _install_facade() -> None:
    """Вмикає передачу підмін. Тести вантажать цей файл через `spec_from_file_location`
    без запису в `sys.modules`, тому сам модуль шукаємо за його словником."""
    module = sys.modules.get(__name__)
    if module is None or module.__dict__ is not globals():
        module = next(
            (
                referrer
                for referrer in _gc.get_referrers(globals())
                if isinstance(referrer, _types.ModuleType) and referrer.__dict__ is globals()
            ),
            None,
        )
    if module is not None and type(module) is _types.ModuleType:
        module.__class__ = _FacadeModule
        # Свіже завантаження фасаду — свіжі (tkinter) значення в пакеті, як
        # було, коли кожне завантаження файлу мало власні глобальні імена.
        for name in ("tk", "messagebox", "filedialog", "DISABLED", "NORMAL", "W"):
            if name in globals():
                setattr(module, name, globals()[name])


_install_facade()


if __name__ == "__main__":
    # Tk-версії більше немає: запуск цього файлу відкриває Qt-оболонку.
    from generate_extracts_qt import main

    raise SystemExit(main())
