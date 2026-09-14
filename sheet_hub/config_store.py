from __future__ import annotations

import json
import os
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from .models import SourceConfig


APP_NAME = "SheetDataHub"
DEFAULT_FIELDS = {
    "日期": ["日期", "时间", "date", "datetime", "创建时间"],
    "名字": ["名字", "姓名", "名称", "name"],
    "链接": ["链接", "网址", "url", "link"],
    "来源": ["来源", "渠道", "平台", "source"],
    "号码": ["号码", "手机号", "手机号码", "电话", "联系电话", "phone", "number"],
}


def default_data_dir() -> Path:
    root = Path(os.getenv("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return root / APP_NAME


class ConfigStore:
    def __init__(self, data_dir: str | Path | None = None):
        self.data_dir = Path(data_dir) if data_dir else default_data_dir()
        self.data_dir.mkdir(parents=True, exist_ok=True)
        (self.data_dir / "databases").mkdir(exist_ok=True)
        self.db_path = self.data_dir / "app.sqlite"
        self._initialize()

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _initialize(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS sources (
                    id TEXT PRIMARY KEY,
                    data TEXT NOT NULL
                );
                CREATE TABLE IF NOT EXISTS logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                    level TEXT NOT NULL,
                    operation TEXT NOT NULL,
                    message TEXT NOT NULL,
                    detail TEXT NOT NULL DEFAULT ''
                );
                CREATE TABLE IF NOT EXISTS extracted (
                    dedup_key TEXT PRIMARY KEY,
                    extracted_at TEXT NOT NULL DEFAULT (datetime('now','localtime')),
                    destination TEXT NOT NULL
                );
                """
            )
        defaults = {
            "field_aliases": DEFAULT_FIELDS,
            "global_excludes": ["说明", "统计", "模板", "配置"],
            "write_aggregate": True,
            "max_rows_per_db": 500000,
            "credential_path": "",
            "column_schema_enabled": False,
            "column_schema": [],
            "google_output_url": "",
            "google_output_sheet": "提取结果",
            "extract_destination_type": "local",
            "extract_output_path": "",
            "query_source": "extract",
            "query_field": "号码",
            "query_fields_by_mode": {},
            "query_direct_source_id": "",
            "query_exact": True,
            "query_fuzzy": False,
            "query_date_enabled": False,
            "query_date_field": "日期",
            "query_start_date": "",
            "query_end_date": "",
            "extract_mode": "aggregate",
            "extract_direct_source_id": "",
            "extract_dedup_fields": "号码,日期",
            "extract_date_field": "日期",
            "extract_column_schema_enabled": False,
            "extract_column_schema": [],
        }
        for key, value in defaults.items():
            if self.get(key, None) is None:
                self.set(key, value)
        # Migrate existing installations: early versions omitted the common
        # “手机号码” header, which made the configured 号码 query field empty.
        aliases = self.get("field_aliases", DEFAULT_FIELDS)
        phone_aliases = aliases.setdefault("号码", [])
        if not any(str(name).strip().casefold() == "手机号码" for name in phone_aliases):
            phone_aliases.append("手机号码")
            self.set("field_aliases", aliases)

    def get(self, key: str, default: Any = None) -> Any:
        with self._connect() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
        if not row:
            return default
        try:
            return json.loads(row["value"])
        except json.JSONDecodeError:
            return row["value"]

    def set(self, key: str, value: Any) -> None:
        encoded = json.dumps(value, ensure_ascii=False)
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO settings(key,value) VALUES(?,?) "
                "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, encoded),
            )

    def load_sources(self) -> list[SourceConfig]:
        with self._connect() as conn:
            rows = conn.execute("SELECT data FROM sources ORDER BY rowid").fetchall()
        return [SourceConfig.from_dict(json.loads(row["data"])) for row in rows]

    def save_source(self, source: SourceConfig) -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO sources(id,data) VALUES(?,?) "
                "ON CONFLICT(id) DO UPDATE SET data=excluded.data",
                (source.id, json.dumps(source.to_dict(), ensure_ascii=False)),
            )

    def delete_source(self, source_id: str) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM sources WHERE id=?", (source_id,))

    def log(self, level: str, operation: str, message: str, detail: str = "") -> None:
        with self._connect() as conn:
            conn.execute(
                "INSERT INTO logs(level,operation,message,detail) VALUES(?,?,?,?)",
                (level, operation, message, detail),
            )

    def read_logs(self, limit: int = 1000) -> list[sqlite3.Row]:
        with self._connect() as conn:
            return conn.execute(
                "SELECT created_at,level,operation,message,detail FROM logs "
                "ORDER BY id DESC LIMIT ?", (limit,)
            ).fetchall()

    def clear_logs(self) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM logs")

    def was_extracted(self, dedup_key: str) -> bool:
        with self._connect() as conn:
            return conn.execute(
                "SELECT 1 FROM extracted WHERE dedup_key=?", (dedup_key,)
            ).fetchone() is not None

    def mark_extracted(self, keys: list[str], destination: str) -> None:
        with self._connect() as conn:
            conn.executemany(
                "INSERT OR IGNORE INTO extracted(dedup_key,destination) VALUES(?,?)",
                [(key, destination) for key in keys],
            )
