from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True, slots=True)
class DocumentVariant:
    name: str
    fields: tuple[tuple[str, Any], ...] = ()

    def values(self) -> dict[str, Any]:
        return {"name": self.name, **dict(self.fields)}


@dataclass(frozen=True, slots=True)
class DocumentOperation:
    kind: str
    parameters: tuple[tuple[str, Any], ...] = ()

    def options(self) -> dict[str, Any]:
        return dict(self.parameters)


@dataclass(frozen=True, slots=True)
class WordDocumentBatch:
    source_path: str
    variants: tuple[DocumentVariant, ...]
    operations: tuple[DocumentOperation, ...] = ()

    def with_operation(self, kind: str, **parameters: Any) -> WordDocumentBatch:
        operation = DocumentOperation(kind, tuple(parameters.items()))
        return WordDocumentBatch(
            source_path=self.source_path,
            variants=self.variants,
            operations=(*self.operations, operation),
        )

    def __len__(self) -> int:
        return len(self.variants)

    def __str__(self) -> str:
        return f"{len(self.variants)} документів · {len(self.operations)} операцій"
