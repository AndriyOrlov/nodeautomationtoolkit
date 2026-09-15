"""Порожні білі зображення в змісті, перенесеному з наказу.

У наказах із систем електронного документообігу трапляються зображення, які
нічого не показують: білий прямокутник на місці печатки чи QR-коду, прозора
заглушка. У витяг, примірник і повідомлення їх не переносимо — вони лише
займають місце й зсувають верстку.

Порожнім вважається зображення, в якому КОЖЕН піксель майже білий або майже
прозорий. Зображення, яке неможливо розібрати (EMF/WMF, пошкоджене) або
зіставити з фігурою, порожнім НЕ вважається: краще лишити зайве, ніж загубити
справжнє. Автофігури й написи не чіпаються — це не зображення, і біла фігура
в наказі може щось закривати.
"""

from __future__ import annotations

import base64
import html
import io
import re

#: Канал світліший за це (0–255) вважаємо білим. Для JPEG лишає запас на шум
#: стиснення, але світло-жовтий фон (синій канал ~200) вже не «білий».
WHITE_LEVEL = 245
#: Піксель прозоріший за це (0–255) вважаємо невидимим.
ALPHA_LEVEL = 16

_PART_RE = re.compile(
    r"<pkg:part\b(?P<attrs>[^>]*)>\s*"
    r"(?:<pkg:binaryData>(?P<data>[^<]*)</pkg:binaryData>|<pkg:xmlData>(?P<xml>.*?)</pkg:xmlData>)"
    r"\s*</pkg:part>",
    re.DOTALL,
)
_PART_NAME_RE = re.compile(r'\bpkg:name="([^"]+)"')
_RELATIONSHIP_RE = re.compile(r"<Relationship\b([^>]*?)/?>")
_ATTRIBUTE_RE = re.compile(r'([\w:]+)="([^"]*)"')
_ANCHOR_RE = re.compile(r"<wp:anchor\b.*?</wp:anchor>", re.DOTALL)
_DOC_PR_NAME_RE = re.compile(r'<wp:docPr\b[^>]*\bname="([^"]*)"')
_EMBED_RE = re.compile(r'\br:embed="([^"]+)"')

_MEDIA_PREFIX = "/word/media/"
_DOCUMENT_PART = "/word/document.xml"
_DOCUMENT_RELS = "/word/_rels/document.xml.rels"

_WD_INLINE_PICTURE_TYPES = {3, 4}  # wdInlineShapePicture, wdInlineShapeLinkedPicture
_MSO_PICTURE_TYPES = {11, 13}  # msoLinkedPicture, msoPicture
_WD_WITHIN_TABLE = 12


def is_blank_image_bytes(data: bytes) -> bool | None:
    """`True` — порожнє, `False` — щось видно, `None` — не вдалося розібрати."""
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as image:
            image.load()
            rgba = image.convert("RGBA")
    except Exception:
        return None
    _alpha_min, alpha_max = rgba.getchannel("A").getextrema()
    if alpha_max < ALPHA_LEVEL:
        return True
    background = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
    flattened = Image.alpha_composite(background, rgba).convert("RGB")
    return all(channel_min >= WHITE_LEVEL for channel_min, _max in flattened.getextrema())


def _package_parts(xml: str) -> dict[str, bytes | str]:
    """Частини плаского OPC (`Range.WordOpenXML`): двійкові — bytes, XML — рядок."""
    parts: dict[str, bytes | str] = {}
    for match in _PART_RE.finditer(xml or ""):
        name_match = _PART_NAME_RE.search(match.group("attrs"))
        if not name_match:
            continue
        name = name_match.group(1)
        if match.group("data") is not None:
            try:
                parts[name] = base64.b64decode("".join(match.group("data").split()), validate=True)
            except Exception:
                parts[name] = b""  # нерозбірне — is_blank_image_bytes віддасть None
        else:
            parts[name] = match.group("xml") or ""
    return parts


def images_in_word_open_xml(xml: str) -> list[bytes]:
    """Усі зображення з `Range.WordOpenXML` у порядку пакета."""
    return [
        data
        for name, data in _package_parts(xml).items()
        if name.startswith(_MEDIA_PREFIX) and isinstance(data, bytes)
    ]


def all_images_blank(xml: str) -> bool:
    """Чи є в діапазоні зображення і чи ВСІ вони порожні."""
    images = images_in_word_open_xml(xml)
    return bool(images) and all(is_blank_image_bytes(data) is True for data in images)


def images_of_floating_picture(xml: str, shape_name: str) -> list[bytes]:
    """Зображення саме тієї плаваючої картинки, що зветься `shape_name`.

    У XML абзацу прив'язки Word віддає ВСІ плаваючі картинки поруч, тож
    «усе в абзаці порожнє» не спрацьовувало, щойно біля білої заглушки стояла
    справжня картинка. Фігуру знаходимо за `wp:docPr name` (це `Shape.Name`),
    далі `r:embed` → зв'язок → файл у `/word/media/`. Якщо хоч щось не
    зіставилось, повертаємо порожній список — і фігуру не чіпаємо. Однакові
    назви зустрічаються; тоді враховуються ВСІ такі фігури.
    """
    parts = _package_parts(xml)
    document = parts.get(_DOCUMENT_PART)
    relationships = parts.get(_DOCUMENT_RELS)
    if not isinstance(document, str) or not isinstance(relationships, str):
        return []
    targets = {}
    for match in _RELATIONSHIP_RE.finditer(relationships):
        attributes = dict(_ATTRIBUTE_RE.findall(match.group(1)))
        if "Id" in attributes and "Target" in attributes and attributes.get("TargetMode") != "External":
            targets[attributes["Id"]] = attributes["Target"]

    images: list[bytes] = []
    for anchor in _ANCHOR_RE.findall(document):
        name_match = _DOC_PR_NAME_RE.search(anchor)
        if not name_match or html.unescape(name_match.group(1)) != shape_name:
            continue
        embeds = _EMBED_RE.findall(anchor)
        if not embeds:
            return []  # фігура без вбудованого зображення (зв'язана чи векторна)
        for relationship_id in embeds:
            target = targets.get(relationship_id, "")
            part_name = target if target.startswith("/") else "/word/" + target
            data = parts.get(part_name) if target else None
            if not isinstance(data, bytes):
                return []
            images.append(data)
    return images


def floating_picture_is_blank(xml: str, shape_name: str) -> bool:
    images = images_of_floating_picture(xml, shape_name)
    return bool(images) and all(is_blank_image_bytes(data) is True for data in images)


def _drop_emptied_paragraph(doc, paragraph) -> None:
    """Прибирає абзац, який тримав лише прибране зображення.

    Викликається ПІСЛЯ видалення зображення: якщо абзац став порожнім, він
    існував тільки заради нього — інакше лишився б зайвий порожній рядок.
    Порожні абзаци, які були в наказі й раніше, сюди не потрапляють.
    """
    try:
        paragraph_range = paragraph.Range
        if (paragraph_range.Text or "").strip("\r\x07\x0b\x0c \t"):
            return
        if paragraph_range.Information(_WD_WITHIN_TABLE):
            return  # знак абзацу в комірці Word не видаляє
        if paragraph_range.InlineShapes.Count or paragraph_range.ShapeRange.Count:
            return
        length_before = doc.Content.End
        paragraph_range.Delete()
        if doc.Content.End >= length_before:
            return  # останній абзац документа Word мовчки не видаляє
    except Exception:
        return


def remove_blank_images(doc, content_range) -> int:
    """Прибирає порожні білі зображення з діапазону. Повертає їх кількість.

    `content_range` — саме Range Word, а не числа: після видалення межі
    змісту зсуваються, а Range відстежує їх сам. Межі після виклику треба
    перечитати з нього ж.
    """
    removed = 0
    try:
        inline_count = int(content_range.InlineShapes.Count)
    except Exception:
        inline_count = 0
    # З кінця: видалення не перенумеровує ще не перевірені зображення.
    for index in range(inline_count, 0, -1):
        try:
            shape = content_range.InlineShapes(index)
            if int(shape.Type) not in _WD_INLINE_PICTURE_TYPES:
                continue
            # XML вбудованої картинки містить лише її саму.
            if not all_images_blank(shape.Range.WordOpenXML):
                continue
            paragraph = shape.Range.Paragraphs(1)
            shape.Delete()
            removed += 1
            _drop_emptied_paragraph(doc, paragraph)
        except Exception:
            continue

    try:
        floating = content_range.ShapeRange
        shapes = [floating.Item(index) for index in range(1, int(floating.Count) + 1)]
    except Exception:
        shapes = []
    for shape in shapes:
        try:
            if int(shape.Type) not in _MSO_PICTURE_TYPES:
                continue
            anchor = shape.Anchor.Paragraphs(1)
            if not floating_picture_is_blank(anchor.Range.WordOpenXML, str(shape.Name)):
                continue
            shape.Delete()
            removed += 1
            _drop_emptied_paragraph(doc, anchor)
        except Exception:
            continue
    return removed
