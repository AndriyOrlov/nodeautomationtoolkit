from __future__ import annotations

import json
from pathlib import Path

from .models import GraphModel


def save_graph(graph: GraphModel, path: Path) -> None:
    path.write_text(
        json.dumps(graph.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_graph(path: Path) -> GraphModel:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Некоректний формат сценарію")
    return GraphModel.from_dict(data)

