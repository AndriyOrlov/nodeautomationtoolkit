"""Тести для модуля порівняння документів (Compare Mode)."""

import io
import zipfile
from pathlib import Path

from nodeautomationtoolkit.builtin_nodes.compare_documents import (
    compare_docx_documents,
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


_BLANK = "<w:p/>"
_ITEM_1 = '<w:p><w:r><w:t>1. Старшого лейтенанта призначити до військової частини А1111.</w:t></w:r></w:p>'
_ITEM_2 = '<w:p><w:r><w:t>2. Капітана призначити до військової частини А2222.</w:t></w:r></w:p>'
_SIGNER = '<w:p><w:r><w:t>Командир військової частини А0001</w:t></w:r></w:p>'


def _pair(tmp_path: Path, reference: list[str], checked: list[str]) -> tuple[Path, Path]:
    ref_path = tmp_path / "ref.docx"
    gen_path = tmp_path / "gen.docx"
    ref_path.write_bytes(_create_test_docx(reference))
    gen_path.write_bytes(_create_test_docx(checked))
    return ref_path, gen_path


def test_blank_paragraphs_are_ignored_by_default(tmp_path: Path):
    """Зайві чи відсутні Enter у перевірюваному документі — не розбіжність."""
    ref_path, gen_path = _pair(
        tmp_path,
        [_ITEM_1, _BLANK, _ITEM_2, _BLANK, _BLANK, _SIGNER],
        [_BLANK, _ITEM_1, _BLANK, _BLANK, _BLANK, _ITEM_2, _SIGNER, _BLANK],
    )
    res = compare_docx_documents(ref_path, gen_path, mode="витяги")
    assert res.is_identical is True
    assert res.discrepancies == []
    assert all(row["status"] == "EQUAL" for row in res.side_by_side_rows)
    assert len(res.side_by_side_rows) == 3


def test_blank_paragraphs_do_not_hide_real_differences(tmp_path: Path):
    ref_path, gen_path = _pair(
        tmp_path,
        [_ITEM_1, _BLANK, _ITEM_2, _BLANK, _BLANK, _SIGNER],
        [_ITEM_1, _BLANK, _BLANK, _SIGNER],
    )
    res = compare_docx_documents(ref_path, gen_path, mode="витяги")
    assert not res.is_identical
    assert [d.item_label for d in res.discrepancies] == ["Пункт 2."]


def test_blank_paragraphs_can_still_be_compared_when_asked(tmp_path: Path):
    ref_path, gen_path = _pair(
        tmp_path,
        [_ITEM_1, _BLANK, _ITEM_2, _BLANK, _BLANK, _SIGNER],
        [_ITEM_1, _ITEM_2, _SIGNER],
    )
    res = compare_docx_documents(ref_path, gen_path, mode="витяги", ignore_blank_paragraphs=False)
    issue_types = {d.issue_type for d in res.discrepancies}
    assert "Відсутній ентер перед пунктом" in issue_types
    assert "Відступ перед підписантом" in issue_types
