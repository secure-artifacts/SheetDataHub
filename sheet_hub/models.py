from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class SourceConfig:
    id: str
    name: str
    url: str
    include_sheets: list[str] = field(default_factory=list)
    exclude_sheets: list[str] = field(default_factory=list)
    header_row: int = 1
    enabled: bool = True
    credential_path: str = ""
    column_schema_enabled: bool = False
    column_schema: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "SourceConfig":
        allowed = {item.name for item in cls.__dataclass_fields__.values()}
        return cls(**{key: val for key, val in value.items() if key in allowed})


@dataclass
class Record:
    source_id: str
    source_name: str
    spreadsheet_id: str
    sheet_name: str
    row_number: int
    values: dict[str, str]
    row_hash: str

    def display(self) -> dict[str, str]:
        result = dict(self.values)
        result.update(
            {
                "数据源": self.source_name,
                "子Sheet": self.sheet_name,
                "原始行号": str(self.row_number),
            }
        )
        return result

