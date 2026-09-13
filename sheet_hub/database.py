from __future__ import annotations

import json
import sqlite3
from pathlib import Path

from .models import Record


class AggregateDatabase:
    def __init__(self, directory: str | Path, max_rows: int = 500000):
        self.directory = Path(directory)
        self.directory.mkdir(parents=True, exist_ok=True)
        self.max_rows = max(1000, int(max_rows))

    def paths(self) -> list[Path]:
        return sorted(
            path for path in self.directory.glob("aggregate_*.sqlite")
            if path.stem.removeprefix("aggregate_").isdigit()
        )

    def _path(self, number: int) -> Path:
        return self.directory / f"aggregate_{number:03d}.sqlite"

    @staticmethod
    def _init(conn: sqlite3.Connection) -> None:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS records (
                row_hash TEXT PRIMARY KEY,
                source_id TEXT NOT NULL,
                source_name TEXT NOT NULL,
                spreadsheet_id TEXT NOT NULL,
                sheet_name TEXT NOT NULL,
                row_number INTEGER NOT NULL,
                payload TEXT NOT NULL,
                updated_at TEXT NOT NULL DEFAULT (datetime('now','localtime'))
            );
            CREATE INDEX IF NOT EXISTS ix_records_source ON records(source_id, sheet_name);
            """
        )

    def replace_all(self, records: list[Record]) -> tuple[int, int]:
        temp_paths: list[Path] = []
        try:
            chunks = [records[i:i + self.max_rows] for i in range(0, len(records), self.max_rows)] or [[]]
            for number, chunk in enumerate(chunks, start=1):
                temp = self.directory / f"aggregate_{number:03d}.tmp.sqlite"
                if temp.exists():
                    temp.unlink()
                conn = sqlite3.connect(temp)
                try:
                    self._init(conn)
                    conn.executemany(
                        "INSERT OR REPLACE INTO records(row_hash,source_id,source_name,spreadsheet_id,sheet_name,row_number,payload) VALUES(?,?,?,?,?,?,?)",
                        [
                            (r.row_hash, r.source_id, r.source_name, r.spreadsheet_id, r.sheet_name, r.row_number,
                             json.dumps(r.values, ensure_ascii=False))
                            for r in chunk
                        ],
                    )
                    conn.commit()
                finally:
                    conn.close()
                temp_paths.append(temp)
            for old in self.paths():
                old.unlink()
            for number, temp in enumerate(temp_paths, start=1):
                temp.replace(self._path(number))
            return len(records), len(temp_paths)
        except Exception:
            for temp in temp_paths:
                if temp.exists():
                    temp.unlink()
            raise

    def all_records(self) -> list[Record]:
        result: list[Record] = []
        for path in self.paths():
            conn = sqlite3.connect(path)
            try:
                rows = conn.execute(
                    "SELECT source_id,source_name,spreadsheet_id,sheet_name,row_number,payload,row_hash FROM records"
                ).fetchall()
            finally:
                conn.close()
            result.extend(Record(*row[:5], json.loads(row[5]), row[6]) for row in rows)
        return result

    def query(self, field: str, value: str, exact: bool = True) -> list[Record]:
        needle = value.strip().casefold()
        result: list[Record] = []
        for record in self.all_records():
            haystack = record.values.get(field, "").strip().casefold()
            if (haystack == needle) if exact else (needle in haystack):
                result.append(record)
        return result
