from __future__ import annotations

from typing import Literal

from pydantic import BaseModel, Field

from nodeautomationtoolkit.core.definition import node
from nodeautomationtoolkit.core.embedded_llm import MODEL_ALIAS, SERVER_API_KEY, embedded_base_url
from nodeautomationtoolkit.core.local_llm import LocalLlmClient, LocalLlmConfig, LocalLlmProvider
from nodeautomationtoolkit.core.word_types import WordParagraphs


class AiOrderBlock(BaseModel):
    block_type: Literal[
        "header",
        "section",
        "reason",
        "action",
        "item",
        "signature",
        "certification",
        "other",
    ]
    paragraph_indices: list[int]
    section_block: int = -1
    reason_block: int = -1
    action_block: int = -1
    recipients: list[str] = Field(default_factory=list)
    confidence: float = Field(ge=0.0, le=1.0)
    note: str = ""


class AiAmbiguity(BaseModel):
    paragraph_indices: list[int]
    question: str


class AiOrderStructure(BaseModel):
    blocks: list[AiOrderBlock]
    ambiguities: list[AiAmbiguity] = Field(default_factory=list)


STRUCTURE_PROMPT = """Ти локальний аналізатор структури українського наказу.
Поверни лише JSON за схемою. Не переписуй, не виправляй і не вигадуй текст.

Кожен непорожній абзац має потрапити рівно в один блок через paragraph_indices.
Суміжні абзаци одного багатoабзацного пункту об'єднуй в один блок item.

Типи:
- header: шапка наказу до основної частини;
- section: параграф або розділ;
- reason: причина чи підстава для наступних пунктів;
- action: шапка дії, наприклад ЗВІЛЬНИТИ, ПРИЗНАЧИТИ, ПЕРЕМІСТИТИ;
- item: нумерований пункт разом з усіма його продовженнями та підпунктами;
- signature: посада, ім'я та підпис особи наприкінці наказу;
- certification: блок «З оригіналом згідно/вірно» та особа, яка засвідчує;
- other: текст, роль якого неможливо надійно визначити.

section_block, reason_block та action_block містять номер батьківського блока у
вихідному масиві або -1. recipients можуть містити лише значення з переданого списку.
Відправник із section/reason/action успадковується наступними пунктами до нового
відповідного блока. Якщо пункт має двох відправників, вкажи обох.
confidence нижче 0.75 означає, що місце треба показати людині на перевірку.
"""


def _paragraph_texts(text: str, paragraphs: WordParagraphs | None) -> list[str]:
    if paragraphs is not None:
        return [item.text for item in paragraphs if item.text.strip()]
    return [line.strip() for line in text.splitlines() if line.strip()]


def _known_markers(markers: list | None, markers_text: str) -> list[str]:
    values = [str(item).strip() for item in (markers or []) if str(item).strip()]
    if not values:
        values = [line.strip() for line in markers_text.splitlines() if line.strip()]
    return list(dict.fromkeys(values))


def _normalize_structure(
    structure: AiOrderStructure,
    paragraphs: list[str],
    markers: list[str],
) -> dict:
    marker_lookup = {marker.casefold(): marker for marker in markers}
    used: set[int] = set()
    resolved: list[dict] = []
    for source in structure.blocks:
        indices = sorted(
            {
                index
                for index in source.paragraph_indices
                if 0 <= index < len(paragraphs) and index not in used
            }
        )
        if not indices:
            continue
        used.update(indices)
        recipients = [
            marker_lookup[value.casefold()]
            for value in source.recipients
            if value.casefold() in marker_lookup
        ]
        resolved.append(
            {
                "type": source.block_type,
                "paragraph_indices": indices,
                "text": "\n".join(paragraphs[index] for index in indices),
                "section_block": source.section_block,
                "reason_block": source.reason_block,
                "action_block": source.action_block,
                "recipients": list(dict.fromkeys(recipients)),
                "confidence": source.confidence,
                "note": source.note,
            }
        )
    for index, value in enumerate(paragraphs):
        if index not in used:
            resolved.append(
                {
                    "type": "other",
                    "paragraph_indices": [index],
                    "text": value,
                    "section_block": -1,
                    "reason_block": -1,
                    "action_block": -1,
                    "recipients": [],
                    "confidence": 0.0,
                    "note": "Модель пропустила абзац",
                }
            )
    resolved.sort(key=lambda item: item["paragraph_indices"][0])
    for index, block in enumerate(resolved):
        block["index"] = index
    ambiguities = [item.model_dump() for item in structure.ambiguities]
    ambiguities.extend(
        {
            "paragraph_indices": block["paragraph_indices"],
            "question": block["note"] or "Перевірте тип і батьківський блок",
        }
        for block in resolved
        if block["confidence"] < 0.75
    )
    items = [block["text"] for block in resolved if block["type"] == "item"]
    summary = (
        f"Абзаців: {len(paragraphs)} · блоків: {len(resolved)} · "
        f"пунктів: {len(items)} · неоднозначностей: {len(ambiguities)}"
    )
    return {
        "blocks": resolved,
        "items": items,
        "ambiguities": ambiguities,
        "summary": summary,
    }


@node(
    name="AI: розібрати структуру наказу",
    category="Наказ · Локальна AI",
    description=(
        "Локальна Qwen3 4B визначає параграфи, причини, шапки дій, цілі "
        "багатоабзацні пункти, підпис і засвідчення. Документ не виходить з ПК."
    ),
    type_id="builtin.order.ai_structure",
    outputs={
        "blocks": "List",
        "items": "List",
        "ambiguities": "List",
        "summary": "str",
    },
    preview_policy="never",
)
def analyze_order_structure_locally(
    text: str = "",
    paragraphs: WordParagraphs | None = None,
    markers: list | None = None,
    markers_text: str = "",
) -> dict:
    values = _paragraph_texts(text, paragraphs)
    if not values:
        raise ValueError("У документі немає тексту для аналізу")
    known = _known_markers(markers, markers_text)
    numbered = "\n".join(f"[{index}] {value}" for index, value in enumerate(values))
    prompt = f"Відомі відправники: {known or ['не задані']}\n\nАбзаци наказу:\n{numbered}"
    client = LocalLlmClient(
        LocalLlmConfig(
            provider=LocalLlmProvider.EMBEDDED,
            base_url=embedded_base_url(),
            model=MODEL_ALIAS,
            api_key=SERVER_API_KEY,
            timeout_seconds=300,
        )
    )
    payload = client.generate_structured(
        system_prompt=STRUCTURE_PROMPT,
        user_prompt=prompt,
        schema_name="order_structure",
        schema=AiOrderStructure.model_json_schema(),
    )
    try:
        structure = AiOrderStructure.model_validate(payload)
    except (TypeError, ValueError) as error:
        raise RuntimeError("Локальна модель повернула некоректну структуру") from error
    return _normalize_structure(structure, values, known)
