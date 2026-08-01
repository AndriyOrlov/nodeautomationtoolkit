from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True, slots=True)
class DataTable:
    columns: tuple[str, ...]
    rows: tuple[tuple[Any, ...], ...] = ()
    title: str = "Таблиця"

    def to_csv(self, delimiter: str = ";") -> str:
        stream = io.StringIO(newline="")
        writer = csv.writer(stream, delimiter=delimiter, lineterminator="\n")
        writer.writerow(self.columns)
        writer.writerows(self.rows)
        return stream.getvalue()

    def as_dicts(self) -> list[dict[str, Any]]:
        return [dict(zip(self.columns, row, strict=False)) for row in self.rows]
