"""Тести для модуля порівняння документів (Compare Mode)."""

import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

from nodeautomationtoolkit.builtin_nodes.compare_documents import (
    DocParagraph,
    compare_docx_documents,
    generate_ai_chat_report,
    _parse_docx_paragraphs,
)


def _create_test_docx(paragraphs_xml: list[str]) -> bytes:
    """Створює валідний in-memory DOCX файл з заданим списком XML-абзаців."""
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">\n'
        '  <w:body>\n'
        + "\n".join(paragraphs_xml) +
        '\n  </w:body>\n'
        '</w:document>'
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>\n'
        '</Relationships>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
        '  <Default Extension="xml" ContentType="application/xml"/>\n'
        '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>\n'
        '</Types>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", doc_xml)
    return buf.getvalue()


def test_compare_identical_documents(tmp_path: Path):
    p_xml = [
        '<w:p><w:pPr><w:jc w:val="both"/><w:ind w:firstLine="709"/></w:pPr><w:r><w:t>1. Старшого лейтенанта призначити до військової частини А1111.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>1995 р.н., освіта вища</w:t></w:r></w:p>',
    ]
    docx_bytes = _create_test_docx(p_xml)

    ref_path = tmp_path / "ref.docx"
    gen_path = tmp_path / "gen.docx"
    ref_path.write_bytes(docx_bytes)
    gen_path.write_bytes(docx_bytes)

    res = compare_docx_documents(ref_path, gen_path, mode="витяги")
    assert res.is_identical is True
    assert len(res.discrepancies) == 0
    assert "повністю відповідають усім правилам" in res.summary_text


def test_compare_cipher_mismatch(tmp_path: Path):
    p_ref = [
        '<w:p><w:r><w:t>1. Направити до військової частини А1111.</w:t></w:r></w:p>',
    ]
    p_gen = [
        '<w:p><w:r><w:t>1. Направити до військової частини А2222.</w:t></w:r></w:p>',
    ]
    ref_path = tmp_path / "ref.docx"
    gen_path = tmp_path / "gen.docx"
    ref_path.write_bytes(_create_test_docx(p_ref))
    gen_path.write_bytes(_create_test_docx(p_gen))

    res = compare_docx_documents(ref_path, gen_path, mode="витяги")
    assert res.is_identical is False
    assert len(res.discrepancies) >= 1
    d = res.discrepancies[0]
    assert d.issue_type == "Підстановка шифру ВЧ (Cipher)"
    assert "💡" in res.ai_chat_report


def test_compare_alignment_and_indent_mismatch(tmp_path: Path):
    p_ref = [
        '<w:p><w:pPr><w:jc w:val="both"/><w:ind w:firstLine="709"/></w:pPr><w:r><w:t>11. Капітана Іванова зарахувати.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>1990 р.н., освіта повна</w:t></w:r></w:p>',
    ]
    p_gen = [
        '<w:p><w:pPr><w:jc w:val="left"/><w:ind w:firstLine="0"/></w:pPr><w:r><w:t>11. Капітана Іванова зарахувати.</w:t></w:r></w:p>',
        '<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>1990 р.н., освіта повна</w:t></w:r></w:p>',
    ]
    ref_path = tmp_path / "ref.docx"
    gen_path = tmp_path / "gen.docx"
    ref_path.write_bytes(_create_test_docx(p_ref))
    gen_path.write_bytes(_create_test_docx(p_gen))

    res = compare_docx_documents(ref_path, gen_path, mode="витяги")
    assert res.is_identical is False
    types = [d.issue_type for d in res.discrepancies]
    assert "Вирівнювання (Alignment)" in types
    assert "Абзацний відступ (Indent)" in types


def test_ai_chat_report_is_confidential_and_rule_only(tmp_path: Path):
    sensitive_name = "Секретний_Офіцер_Петренко_П.П."
    p_ref = [
        f'<w:p><w:r><w:t>1. {sensitive_name} призначити до частини А5000.</w:t></w:r></w:p>',
    ]
    p_gen = [
        f'<w:p><w:r><w:t>1. {sensitive_name} призначити до частини військової частини А5000.</w:t></w:r></w:p>',
    ]
    ref_path = tmp_path / "ref.docx"
    gen_path = tmp_path / "gen.docx"
    ref_path.write_bytes(_create_test_docx(p_ref))
    gen_path.write_bytes(_create_test_docx(p_gen))

    res = compare_docx_documents(ref_path, gen_path, mode="повідомлення")
    report = res.ai_chat_report
    # Гарантія конфіденційності: текст наказу або прізвище НЕ потрапляє у звіт для AI
    assert sensitive_name not in report
    assert "Звіт розбіжностей правил Compare Mode" in report
    assert "Очікуване правило" in report
    assert "Фактично у генераторі" in report
    assert "Що виправити в коді" in report
    assert "Запит до AI" in report
