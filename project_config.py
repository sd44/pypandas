from __future__ import annotations

from dataclasses import asdict, dataclass, field
import json
from pathlib import Path


def _coerce_split_columns(payload: dict) -> list[str]:
    """兼容旧版 split_column (str) 和新版 split_columns (list[str])。"""
    if "split_columns" in payload:
        raw = payload["split_columns"]
        if isinstance(raw, list):
            return raw
        return []
    old = payload.get("split_column", "")
    if old and isinstance(old, str):
        return [old]
    return []


@dataclass
class ColumnRule:
    original_name: str
    display_name: str = ""
    enabled: bool = True


@dataclass
class ProjectConfig:
    source_files: list[str] = field(default_factory=list)
    active_file: str = ""
    sheet_name: str = ""
    header_row: int = 1
    output_dir: str = ""
    output_prefix: str = "result"
    export_format: str = "xlsx"
    split_columns: list[str] = field(default_factory=list)
    split_values: list[str] = field(default_factory=list)
    sort_column: str = ""
    drop_duplicates: bool = False
    remove_empty_rows: bool = True
    column_rules: list[ColumnRule] = field(default_factory=list)

    def sync_columns(self, columns: list[str]) -> None:
        current = {rule.original_name: rule for rule in self.column_rules}
        synced: list[ColumnRule] = []
        for column in columns:
            rule = current.get(column)
            if rule is None:
                synced.append(ColumnRule(original_name=column, display_name=column, enabled=True))
            else:
                if not rule.display_name:
                    rule.display_name = column
                synced.append(rule)
        self.column_rules = synced
        if self.split_columns:
            self.split_columns = [c for c in self.split_columns if c in columns]
        if self.sort_column not in columns:
            self.sort_column = ""

    def to_dict(self) -> dict:
        data = asdict(self)
        data["column_rules"] = [asdict(rule) for rule in self.column_rules]
        return data

    @classmethod
    def from_dict(cls, payload: dict) -> "ProjectConfig":
        rules = [ColumnRule(**rule) for rule in payload.get("column_rules", [])]
        return cls(
            source_files=payload.get("source_files", []),
            active_file=payload.get("active_file", ""),
            sheet_name=payload.get("sheet_name", ""),
            header_row=int(payload.get("header_row", 1) or 1),
            output_dir=payload.get("output_dir", ""),
            output_prefix=payload.get("output_prefix", "result"),
            export_format=payload.get("export_format", "xlsx"),
            split_columns=_coerce_split_columns(payload),
            split_values=payload.get("split_values", []),
            sort_column=payload.get("sort_column", ""),
            drop_duplicates=bool(payload.get("drop_duplicates", False)),
            remove_empty_rows=bool(payload.get("remove_empty_rows", True)),
            column_rules=rules,
        )

    @classmethod
    def load(cls, path: str | Path) -> "ProjectConfig":
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.from_dict(payload)

    def save(self, path: str | Path) -> None:
        Path(path).write_text(
            json.dumps(self.to_dict(), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
