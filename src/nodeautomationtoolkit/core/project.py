from __future__ import annotations

import json
import os
from pathlib import Path
from tempfile import NamedTemporaryFile

from .models import GraphModel


def save_graph(graph: GraphModel, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    content = json.dumps(graph.to_dict(), ensure_ascii=False, indent=2)
    temporary_path: Path | None = None
    try:
        with NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=path.parent,
            prefix=f".{path.name}.",
            suffix=".tmp",
            delete=False,
        ) as temporary:
            temporary.write(content)
            temporary.flush()
            os.fsync(temporary.fileno())
            temporary_path = Path(temporary.name)
        temporary_path.replace(path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink()


def load_graph(path: Path) -> GraphModel:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Некоректний формат сценарію")
    return GraphModel.from_dict(data)
