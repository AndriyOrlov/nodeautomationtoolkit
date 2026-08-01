from nodeautomationtoolkit.builtin_nodes.order_ai import (
    AiOrderBlock,
    AiOrderStructure,
    _normalize_structure,
)
from nodeautomationtoolkit.core.embedded_llm import MODEL_ALIAS
from nodeautomationtoolkit.core.local_llm import LocalLlmProvider


def test_normalizes_ai_blocks_without_rewriting_text():
    paragraphs = [
        "ПАРАГРАФ 1",
        "ЗВІЛЬНИТИ І ПРИЗНАЧИТИ до АЛЬФА",
        "1. Перший абзац пункту",
        "Продовження цього самого пункту",
        "Командир частини",
    ]
    structure = AiOrderStructure(
        blocks=[
            AiOrderBlock(
                block_type="section", paragraph_indices=[0], confidence=0.99
            ),
            AiOrderBlock(
                block_type="action",
                paragraph_indices=[1],
                recipients=["альфа", "НЕВІДОМИЙ"],
                confidence=0.95,
            ),
            AiOrderBlock(
                block_type="item",
                paragraph_indices=[2, 3],
                action_block=1,
                recipients=["АЛЬФА"],
                confidence=0.9,
            ),
            AiOrderBlock(
                block_type="signature", paragraph_indices=[4], confidence=0.7
            ),
        ]
    )

    result = _normalize_structure(structure, paragraphs, ["АЛЬФА"])

    assert result["blocks"][1]["recipients"] == ["АЛЬФА"]
    assert result["items"] == [
        "1. Перший абзац пункту\nПродовження цього самого пункту"
    ]
    assert len(result["ambiguities"]) == 1
    assert "Абзаців: 5" in result["summary"]


def test_adds_omitted_paragraph_as_ambiguous_other_block():
    structure = AiOrderStructure(
        blocks=[
            AiOrderBlock(block_type="header", paragraph_indices=[0], confidence=1.0)
        ]
    )

    result = _normalize_structure(structure, ["Шапка", "Пропущений текст"], [])

    assert result["blocks"][1]["type"] == "other"
    assert result["blocks"][1]["text"] == "Пропущений текст"
    assert result["ambiguities"][0]["paragraph_indices"] == [1]


def test_embedded_provider_is_available():
    assert LocalLlmProvider.EMBEDDED.value == "Вбудована Qwen3 4B"
    assert MODEL_ALIAS == "embedded-qwen3-4b"
