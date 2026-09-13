from __future__ import annotations

import hashlib
import json
import re
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import date, datetime
from pathlib import Path
from typing import Callable

import openpyxl

from .config_store import ConfigStore
from .database import AggregateDatabase
from .models import Record, SourceConfig
from .source_reader import SourceReader, google_retry, spreadsheet_id


ProgressFn = Callable[[str], None]


def parse_date(value: str) -> date | None:
    text = str(value).strip()
    if not text:
        return None
    if "T" in text:
        text = text.split("T", 1)[0]
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


class DataEngine:
    def __init__(self, store: ConfigStore, progress: ProgressFn | None = None):
        self.store = store
        self.progress = progress or (lambda message: None)

    @property
    def database(self) -> AggregateDatabase:
        return AggregateDatabase(
            self.store.data_dir / "databases",
            int(self.store.get("max_rows_per_db", 500000)),
        )

    def _log(self, operation: str, level: str, message: str, detail: str = "") -> None:
        self.store.log(level, operation, message, detail)
        self.progress(message)

    def _reader(self, operation: str) -> SourceReader:
        schema = self.store.get("column_schema", []) if self.store.get("column_schema_enabled", False) else []
        return SourceReader(
            self.store.get("field_aliases", {}),
            self.store.get("global_excludes", []),
            lambda level, message: self._log(operation, level, message),
            schema,
        )

    def read_sources(self, operation: str = "读取数据源") -> list[Record]:
        sources = [source for source in self.store.load_sources() if source.enabled]
        if not sources:
            raise ValueError("没有启用的数据源")
        reader = self._reader(operation)
        records: list[Record] = []
        failures = 0
        for source in sources:
            try:
                if not source.credential_path:
                    source.credential_path = str(self.store.get("credential_path", "")).strip()
                self._log(operation, "INFO", f"开始读取：{source.name}")
                loaded = reader.read(source)
                records.extend(loaded)
                self._log(operation, "INFO", f"完成读取：{source.name}，共 {len(loaded)} 行")
            except Exception as exc:
                failures += 1
                self._log(operation, "ERROR", f"读取失败：{source.name}", repr(exc))
        if failures == len(sources):
            raise RuntimeError("所有数据源均读取失败，请查看日志")
        return records

    def sync(self) -> dict[str, int]:
        operation = "汇总同步"
        self._log(operation, "INFO", "汇总任务开始")
        records = self.read_sources(operation)
        if not self.store.get("write_aggregate", True):
            self._log(operation, "INFO", f"已读取 {len(records)} 行；当前未启用写入汇总库")
            return {"rows": len(records), "databases": 0}
        unique = {record.row_hash: record for record in records}
        row_count, db_count = self.database.replace_all(list(unique.values()))
        self._log(operation, "INFO", f"汇总完成：写入 {row_count} 行，使用 {db_count} 个数据库文件")
        return {"rows": row_count, "databases": db_count}

    def query(self, field: str, value: str, direct: bool = False, exact: bool = True) -> list[Record]:
        operation = "直接查询" if direct else "汇总库查询"
        self._log(operation, "INFO", f"查询字段“{field}”，值“{value}”")
        if direct:
            records = self.read_sources(operation)
            needle = self._match_value(field, value)
            result = [
                record for record in records
                if ((self._match_value(field, record.values.get(field, "")) == needle) if exact
                    else (needle in self._match_value(field, record.values.get(field, ""))))
            ]
        else:
            result = self.database.query(field, value, exact)
        self._log(operation, "INFO", f"查询完成：匹配 {len(result)} 行")
        return result

    def query_many(
        self,
        field: str,
        values: list[str],
        direct: bool = False,
        exact: bool = True,
    ) -> list[tuple[str, Record | None]]:
        operation = "批量直接查询" if direct else "批量汇总库查询"
        queries = list(dict.fromkeys(value.strip() for value in values if value.strip()))
        if not queries:
            raise ValueError("没有可查询的号码")
        self._log(operation, "INFO", f"开始批量查询：{len(queries)} 个值，字段“{field}”")
        records = self.read_sources(operation) if direct else self.database.all_records()
        results: list[tuple[str, Record | None]] = []
        matched_inputs = 0
        for query_value in queries:
            needle = self._match_value(field, query_value)
            matches = []
            for record in records:
                haystack = self._match_value(field, record.values.get(field, ""))
                if (haystack == needle) if exact else (needle in haystack):
                    matches.append(record)
            if matches:
                matched_inputs += 1
                results.extend((query_value, record) for record in matches)
            else:
                results.append((query_value, None))
        self._log(
            operation,
            "INFO",
            f"批量查询完成：输入 {len(queries)} 个，找到 {matched_inputs} 个，未找到 {len(queries) - matched_inputs} 个",
        )
        return results

    @staticmethod
    def _match_value(field: str, value: str) -> str:
        text = str(value).strip().casefold().lstrip("'")
        if field.strip().casefold() not in {"号码", "手机号", "手机号码", "电话", "联系电话", "phone", "number"}:
            return text
        compact = re.sub(r"[\s,，\-()（）]", "", text)
        if re.fullmatch(r"[+-]?\d+(?:\.0+)?", compact):
            return compact.split(".", 1)[0].lstrip("+")
        if re.fullmatch(r"[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+", compact):
            try:
                return format(Decimal(compact), "f").split(".", 1)[0].lstrip("+")
            except InvalidOperation:
                pass
        digits = re.sub(r"\D", "", compact)
        return digits or compact

    def extract(
        self,
        date_field: str,
        start: date,
        end: date,
        output: str | Path,
        dedup_fields: list[str],
        direct: bool = False,
        output_sheet_name: str = "提取结果",
        destination_type: str = "local",
        google_output_url: str = "",
    ) -> dict[str, int]:
        operation = "时间提取"
        if start > end:
            raise ValueError("开始日期不能晚于结束日期")
        records = self.read_sources(operation) if direct else self.database.all_records()
        selected: list[tuple[Record, str]] = []
        skipped_invalid = 0
        skipped_duplicate = 0
        for record in records:
            current = parse_date(record.values.get(date_field, ""))
            if current is None:
                skipped_invalid += 1
                continue
            if not start <= current <= end:
                continue
            parts = [record.values.get(field, "") for field in dedup_fields]
            if not any(parts):
                parts = [record.spreadsheet_id, record.sheet_name, str(record.row_number), record.row_hash]
            key = hashlib.sha256(json.dumps(parts, ensure_ascii=False).encode("utf-8")).hexdigest()
            if self.store.was_extracted(key):
                skipped_duplicate += 1
                continue
            selected.append((record, key))
        selected_records = [record for record, _ in selected]
        if destination_type == "google":
            if not google_output_url.strip():
                raise ValueError("请填写目标 Google 表格链接")
            self._write_google_sheet(google_output_url, output_sheet_name, selected_records, operation)
            destination_label = f"{google_output_url}#{output_sheet_name}"
        else:
            destination = Path(output)
            destination.parent.mkdir(parents=True, exist_ok=True)
            headers = self._preferred_export_headers(selected_records)
            self._write_xlsx(
                destination,
                selected_records,
                output_sheet_name,
                headers,
                self.store.get("field_aliases", {}),
            )
            destination_label = str(destination)
        self.store.mark_extracted([key for _, key in selected], destination_label)
        self._log(
            operation,
            "INFO",
            f"提取完成：写入 {len(selected)} 行，排除重复 {skipped_duplicate} 行，日期无效 {skipped_invalid} 行",
        )
        return {"written": len(selected), "duplicates": skipped_duplicate, "invalid_dates": skipped_invalid}

    def _preferred_export_headers(self, records: list[Record]) -> list[str]:
        if self.store.get("column_schema_enabled", False):
            configured = [
                str(name).strip()
                for name in self.store.get("column_schema", [])
                if str(name).strip()
            ]
            if configured:
                return ["来源", *[name for name in configured if name != "来源"]]
        headers, _ = self._export_rows(records)
        return headers

    @staticmethod
    def _canonical_header(header: str, aliases: dict[str, list[str]]) -> str:
        clean = str(header).strip()
        folded = clean.casefold()
        if folded == "来源":
            return "来源"
        for canonical, candidates in aliases.items():
            if folded in {
                str(name).strip().casefold()
                for name in [canonical, *candidates]
            }:
                return canonical
        return clean

    @classmethod
    def _headers_compatible(
        cls,
        existing: list[str],
        required: list[str],
        aliases: dict[str, list[str]],
    ) -> bool:
        return Counter(cls._canonical_header(name, aliases) for name in existing) == Counter(
            cls._canonical_header(name, aliases) for name in required
        )

    @classmethod
    def _rows_for_headers(
        cls,
        records: list[Record],
        headers: list[str],
        aliases: dict[str, list[str]],
    ) -> list[list[object]]:
        rows: list[list[object]] = []
        for record in records:
            row: list[object] = []
            for header in headers:
                canonical = cls._canonical_header(header, aliases)
                if canonical == "来源":
                    row.append(record.sheet_name)
                else:
                    row.append(record.values.get(canonical, record.values.get(header, "")))
            rows.append(row)
        return rows

    @staticmethod
    def _export_rows(records: list[Record]) -> tuple[list[str], list[list[object]]]:
        fields: list[str] = []
        for record in records:
            for key in record.values:
                if key != "来源" and key not in fields:
                    fields.append(key)
        headers = ["来源", *fields]
        rows = [[record.sheet_name, *[record.values.get(key, "") for key in fields]] for record in records]
        return headers, rows

    @classmethod
    def _write_xlsx(
        cls,
        path: Path,
        records: list[Record],
        sheet_name: str = "提取结果",
        headers: list[str] | None = None,
        aliases: dict[str, list[str]] | None = None,
    ) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        safe_name = "".join("_" if char in "[]:*?/\\" else char for char in sheet_name.strip())[:31]
        sheet.title = safe_name or "提取结果"
        export_headers, default_rows = cls._export_rows(records)
        headers = headers or export_headers
        rows = cls._rows_for_headers(records, headers, aliases or {}) if headers != export_headers or aliases else default_rows
        sheet.append(headers)
        for row in rows:
            sheet.append(row)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for column in sheet.columns:
            letter = column[0].column_letter
            width = min(45, max(10, max(len(str(cell.value or "")) for cell in column) + 2))
            sheet.column_dimensions[letter].width = width
        workbook.save(path)

    def _write_google_sheet(self, url: str, sheet_name: str, records: list[Record], operation: str) -> None:
        sid = spreadsheet_id(url)
        if not sid:
            raise ValueError("无法识别目标 Google 表格链接")
        credential = str(self.store.get("credential_path", "")).strip()
        if not credential:
            credential = next(
                (source.credential_path for source in self.store.load_sources() if source.credential_path),
                "",
            )
        if not credential:
            raise ValueError("写入 Google 表格需要在“设置”中配置服务账号 JSON")
        logger = lambda level, message: self._log(operation, level, message)
        client = SourceReader._gspread_client(credential)
        http = client.http_client
        metadata = google_retry(
            lambda: http.fetch_sheet_metadata(
                sid, params={"fields": "sheets.properties(sheetId,title)"}
            ),
            logger,
        )
        target_name = sheet_name.strip() or "提取结果"
        aliases = self.store.get("field_aliases", {})
        required_headers = self._preferred_export_headers(records)
        sheet_names = {item["properties"]["title"] for item in metadata.get("sheets", [])}
        created = target_name not in sheet_names
        if created:
            google_retry(
                lambda: http.batch_update(
                    sid,
                    {
                        "requests": [{
                            "addSheet": {
                                "properties": {
                                    "title": target_name,
                                    "gridProperties": {
                                        "rowCount": max(1000, len(records) + 10),
                                        "columnCount": max(20, len(required_headers) + 2),
                                    },
                                }
                            }
                        }]
                    },
                ),
                logger,
            )
            existing_header = []
        else:
            escaped_name = target_name.replace("'", "''")
            header_response = google_retry(
                lambda: http.values_get(
                    sid,
                    f"'{escaped_name}'!1:1",
                    params={"majorDimension": "ROWS", "valueRenderOption": "FORMATTED_VALUE"},
                ),
                logger,
            )
            existing_header = (header_response.get("values") or [[]])[0]
            while existing_header and not str(existing_header[-1]).strip():
                existing_header.pop()
        if existing_header and not self._headers_compatible(existing_header, required_headers, aliases):
            raise ValueError(
                f"目标工作表“{target_name}”的表头与提取字段不一致。"
                f"现有：{existing_header}；需要包含：{required_headers}"
            )
        headers = existing_header or required_headers
        rows = self._rows_for_headers(records, headers, aliases)
        payload = ([headers] if not existing_header else []) + rows
        if not payload:
            return
        escaped_name = target_name.replace("'", "''")
        for start in range(0, len(payload), 5000):
            chunk = payload[start:start + 5000]
            google_retry(
                lambda chunk=chunk: http.values_append(
                    sid,
                    f"'{escaped_name}'!A1",
                    params={"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
                    body={"majorDimension": "ROWS", "values": chunk},
                ),
                logger,
            )
            self._log(operation, "INFO", f"已写入 Google 工作表“{target_name}”：{min(start + len(chunk), len(payload))}/{len(payload)} 行")
