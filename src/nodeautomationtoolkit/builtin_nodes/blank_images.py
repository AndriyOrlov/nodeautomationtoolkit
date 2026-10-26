"""Порожні білі зображення й фігури-заглушки в змісті, перенесеному з наказу.

У наказах із систем електронного документообігу трапляються зображення, які
нічого не показують: білий прямокутник на місці печатки чи QR-коду, прозора
заглушка. У витяг, примірник і повідомлення їх не переносимо — вони лише
займають місце й зсувають верстку.

Порожнім вважається зображення, в якому КОЖЕН піксель майже білий або майже
прозорий. Зображення, яке неможливо розібрати (EMF/WMF, пошкоджене) або
зіставити з фігурою, порожнім НЕ вважається: краще лишити зайве, ніж загубити
справжнє.

Графічні фігури (автофігура, вільна фігура, напис) прибираються з 26.10.2026
(вимога користувача: «бувають прямокутники білі як заглушки»), але лише
порожні: без тексту, заливка біла або відсутня, контур білий або відсутній
(`is_blank_shape`). Кольорова фігура чи фігура з текстом лишається. Група —
лише якщо порожні ВСІ її частини.
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
_MSO_GRAPHIC_TYPES = {1, 5, 17}  # msoAutoShape, msoFreeform, msoTextBox
_MSO_GROUP = 6
_WD_WITHIN_TABLE = 12
#: Колір Word (BGR як ціле): канал світліший за WHITE_LEVEL — «білий».
_MSO_TRUE = -1


def _is_white_rgb(value) -> bool:
    try:
        number = int(value)
    except Exception:
        return False
    if number < 0:
        return False  # автоматичний/тематичний колір — не знаємо, який він
    channels = (number & 0xFF, (number >> 8) & 0xFF, (number >> 16) & 0xFF)
    return all(channel >= WHITE_LEVEL for channel in channels)


def _visible(format_object) -> bool:
    return int(format_object.Visible) == _MSO_TRUE


def is_blank_shape(shape) -> bool:
    """Чи фігура Word — порожня біла заглушка (див. опис модуля).

    Будь-яка помилка COM — «не порожня»: краще лишити зайве, ніж загубити справжнє.
    """
    try:
        kind = int(shape.Type)
        if kind == _MSO_GROUP:
            items = shape.GroupItems
            count = int(items.Count)
            return count > 0 and all(is_blank_shape(items.Item(index)) for index in range(1, count + 1))
        if kind not in _MSO_GRAPHIC_TYPES:
            return False
        try:
            frame = shape.TextFrame
            has_text = int(frame.HasText) == _MSO_TRUE
        except Exception:
            has_text = False  # у фігури немає текстової рамки — значить і тексту немає
        if has_text:
            # Порожній напис Word теж «має текст» — знак абзацу. Дивимось сам текст;
            # не вдалося прочитати — вважаємо, що текст є.
            try:
                text = str(frame.TextRange.Text or "")
            except Exception:
                return False
            if text.strip("\r\n\x07\x0b\x0c \t\u00a0"):
                return False
        fill = shape.Fill
        if _visible(fill) and not (
            _is_white_rgb(fill.ForeColor.RGB) or float(fill.Transparency) >= 0.95
        ):
            return False
        line = shape.Line
        if _visible(line) and not _is_white_rgb(line.ForeColor.RGB):
            return False
        return True
    except Exception:
        return False


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
    """Прибирає порожні білі зображення й фігури-заглушки з діапазону. Повертає їх кількість.

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
            kind = int(shape.Type)
            anchor = shape.Anchor.Paragraphs(1)
            if kind in _MSO_PICTURE_TYPES:
                if not floating_picture_is_blank(anchor.Range.WordOpenXML, str(shape.Name)):
                    continue
            elif not is_blank_shape(shape):
                continue
            shape.Delete()
            removed += 1
            _drop_emptied_paragraph(doc, anchor)
        except Exception:
            continue
    return removed
