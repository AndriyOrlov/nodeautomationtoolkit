"""Порожні білі зображення з наказу не переносяться у витяги, примірники й повідомлення.

Тут перевіряється саме рішення «порожнє чи ні» й зіставлення плаваючої
картинки з її файлом: прибирати треба лише те, на чому справді нічого немає,
і ніколи — зображення, яке не вдалося розібрати чи зіставити. Роботу з живим
Word перевіряє `scripts/e2e_copies/check_blank_images_and_fonts.py`.
"""

import base64
import io

import pytest
from PIL import Image, ImageDraw

from nodeautomationtoolkit.builtin_nodes.blank_images import (
    all_images_blank,
    floating_picture_is_blank,
    images_in_word_open_xml,
    images_of_floating_picture,
    is_blank_image_bytes,
)


def _png(color, size=(40, 20), mode="RGBA", draw=None) -> bytes:
    image = Image.new(mode, size, color)
    if draw:
        draw(ImageDraw.Draw(image))
    buffer = io.BytesIO()
    image.save(buffer, format="PNG")
    return buffer.getvalue()


def _jpeg(color, size=(40, 20)) -> bytes:
    buffer = io.BytesIO()
    Image.new("RGB", size, color).save(buffer, format="JPEG", quality=60)
    return buffer.getvalue()


WHITE = _png((255, 255, 255, 255))
DOTTED = _png((255, 255, 255, 255), draw=lambda d: d.point((5, 5), fill=(0, 0, 0, 255)))


def _binary_part(name: str, data: bytes) -> str:
    encoded = base64.b64encode(data).decode()
    wrapped = "\n".join(encoded[i : i + 76] for i in range(0, len(encoded), 76))
    return (
        f'<pkg:part pkg:name="{name}" pkg:contentType="image/png" pkg:compression="store">'
        f"<pkg:binaryData>{wrapped}</pkg:binaryData></pkg:part>"
    )


def _xml_part(name: str, body: str) -> str:
    return f'<pkg:part pkg:name="{name}" pkg:contentType="application/xml"><pkg:xmlData>{body}</pkg:xmlData></pkg:part>'


def _package(*parts: str) -> str:
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' + "".join(parts) + "</pkg:package>"


def _flat_opc(*images: bytes) -> str:
    return _package(
        _xml_part("/word/document.xml", "<w:document/>"),
        *(_binary_part(f"/word/media/image{number}.png", data) for number, data in enumerate(images, 1)),
    )


def _anchor(name: str, embed: str | None) -> str:
    blip = f'<a:blip r:embed="{embed}"/>' if embed else "<a:prstGeom/>"
    return f'<wp:anchor behindDoc="0"><wp:docPr id="7" name="{name}"/><a:graphic>{blip}</a:graphic></wp:anchor>'


def _floating_package(anchors: str, media: dict[str, bytes], relations: dict[str, str]) -> str:
    rels = "".join(
        f'<Relationship Target="{target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Id="{rid}"/>'
        for rid, target in relations.items()
    )
    return _package(
        _xml_part("/word/_rels/document.xml.rels", f"<Relationships>{rels}</Relationships>"),
        _xml_part("/word/document.xml", f"<w:document><w:body><w:p>{anchors}</w:p></w:body></w:document>"),
        *(_binary_part(name, data) for name, data in media.items()),
    )


# ── порожнє чи ні ──────────────────────────────────────────────────────────
@pytest.mark.parametrize(
    "data",
    [WHITE, _png((0, 0, 0, 0)), _png((250, 251, 252, 255)), _jpeg((255, 255, 255)), _png((255, 255, 255), mode="RGB")],
    ids=["білий", "прозорий", "майже білий", "білий JPEG", "білий без альфи"],
)
def test_blank_images_are_recognised(data):
    assert is_blank_image_bytes(data) is True


@pytest.mark.parametrize(
    "data",
    [
        DOTTED,
        _png((255, 255, 200, 255)),
        _png((0, 0, 0, 0), draw=lambda d: d.line((0, 0, 39, 19), fill=(20, 20, 160, 255))),
        _png((255, 255, 255, 255), draw=lambda d: d.rectangle((10, 5, 30, 15), outline=(200, 200, 200, 255))),
    ],
    ids=["одна чорна точка", "світло-жовтий фон", "лінія на прозорому", "сіра рамка"],
)
def test_images_with_anything_visible_are_kept(data):
    assert is_blank_image_bytes(data) is False


@pytest.mark.parametrize("data", [b"", b"not an image", b"\x01\x00\x00\x00 EMF"], ids=["порожньо", "сміття", "EMF"])
def test_unreadable_images_are_never_called_blank(data):
    assert is_blank_image_bytes(data) is None


# ── вбудовані: XML діапазону містить лише свої зображення ─────────────────
def test_images_are_taken_only_from_media_parts_in_order():
    assert images_in_word_open_xml(_flat_opc(WHITE, DOTTED)) == [WHITE, DOTTED]


def test_all_images_blank_needs_every_image_blank():
    assert all_images_blank(_flat_opc(WHITE, _png((0, 0, 0, 0)))) is True
    assert all_images_blank(_flat_opc(WHITE, DOTTED)) is False
    assert all_images_blank(_flat_opc()) is False
    assert all_images_blank("") is False


def test_broken_base64_keeps_the_image():
    xml = _flat_opc(WHITE).replace("<pkg:binaryData>", "<pkg:binaryData>@@@", 1)
    assert all_images_blank(xml) is False


# ── плаваючі: абзац прив'язки віддає ВСІ картинки поруч ────────────────────
NEIGHBOURS = _floating_package(
    _anchor("Рисунок 1", "rId5") + _anchor("Рисунок 2", "rId6"),
    {"/word/media/image1.png": DOTTED, "/word/media/image2.png": WHITE},
    {"rId5": "media/image1.png", "rId6": "media/image2.png"},
)


def test_white_floating_picture_next_to_a_real_one_is_blank():
    """Саме цей випадок живий Word показав: обидві фігури в одному XML."""
    assert all_images_blank(NEIGHBOURS) is False  # старе правило «усе порожнє» тут безсиле
    assert floating_picture_is_blank(NEIGHBOURS, "Рисунок 2") is True
    assert floating_picture_is_blank(NEIGHBOURS, "Рисунок 1") is False


def test_unknown_or_unmatched_floating_picture_is_kept():
    assert floating_picture_is_blank(NEIGHBOURS, "Рисунок 9") is False
    no_embed = _floating_package(_anchor("Фігура", None), {}, {})
    assert floating_picture_is_blank(no_embed, "Фігура") is False
    missing_media = _floating_package(_anchor("Рисунок 3", "rId7"), {}, {"rId7": "media/image7.png"})
    assert images_of_floating_picture(missing_media, "Рисунок 3") == []


def test_duplicate_names_need_every_such_picture_blank():
    duplicates = _floating_package(
        _anchor("Рисунок 1", "rId5") + _anchor("Рисунок 1", "rId6"),
        {"/word/media/image1.png": WHITE, "/word/media/image2.png": DOTTED},
        {"rId5": "media/image1.png", "rId6": "media/image2.png"},
    )
    assert floating_picture_is_blank(duplicates, "Рисунок 1") is False


def test_escaped_names_and_absolute_targets_are_matched():
    package = _floating_package(
        _anchor("Рисунок &quot;печатка&quot;", "rId8"),
        {"/word/media/image8.png": WHITE},
        {"rId8": "/word/media/image8.png"},
    )
    assert floating_picture_is_blank(package, 'Рисунок "печатка"') is True
