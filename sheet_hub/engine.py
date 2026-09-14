from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from collections import Counter
from decimal import Decimal, InvalidOperation
from datetime import date, datetime
from pathlib import Path
from typing import Callable

import openpyxl

from .config_store import ConfigStore
from .database import AggregateDatabase
from .models import Record, SourceConfig
from .source_reader import SourceReader, google_retry, schema_field_names, spreadsheet_id


ProgressFn = Callable[[str], None]


def parse_date(value: str) -> date | None:
    text = str(value).strip()
    if not text:
        return None
    text = re.sub(r"\s*(年|月)\s*", lambda match: "-" if match.group(1) in {"年", "月"} else match.group(0), text)
    text = re.sub(r"\s*日\s*$", "", text)
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
    SEMANTIC_FIELD_HINTS = {
        "日期": ("日期", "date", "datetime"),
        "号码": ("手机", "电话", "号码", "phone", "number", "tel"),
        "名字": ("姓名", "名字", "名称", "name"),
        "链接": ("链接", "网址", "link", "url"),
        "来源": ("来源", "渠道", "平台", "source"),
    }

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

    def _schema_for_source(self, source: SourceConfig | None, use_schema: bool) -> list[str]:
        if not use_schema:
            return []
        if source is not None and source.column_schema_enabled and source.column_schema:
            return list(source.column_schema)
        if self.store.get("column_schema_enabled", False):
            return list(self.store.get("column_schema", []) or [])
        return []

    def _reader(
        self,
        operation: str,
        *,
        source: SourceConfig | None = None,
        use_schema: bool = True,
        schema: list[str] | None = None,
        global_excludes: list[str] | None = None,
    ) -> SourceReader:
        resolved = schema if schema is not None else self._schema_for_source(source, use_schema)
        excludes = self.store.get("global_excludes", []) if global_excludes is None else global_excludes
        return SourceReader(
            self.store.get("field_aliases", {}),
            excludes,
            lambda level, message: self._log(operation, level, message),
            resolved,
        )

    def read_sources(self, operation: str = "读取数据源", source_id: str = "") -> list[Record]:
        sources = [source for source in self.store.load_sources() if source.enabled]
        if source_id.strip():
            sources = [source for source in sources if source.id == source_id.strip()]
        if not sources:
            raise ValueError("没有启用的数据源" if not source_id.strip() else "没有找到所选数据源")
        records: list[Record] = []
        failures = 0
        for source in sources:
            try:
                if not source.credential_path:
                    source.credential_path = str(self.store.get("credential_path", "")).strip()
                self._log(operation, "INFO", f"开始读取：{source.name}")
                loaded = self._reader(operation, source=source).read(source)
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

    def read_extract_table(self, target: str, sheet_name: str = "") -> list[Record]:
        path_or_url = str(target or "").strip()
        if not path_or_url:
            raise ValueError("请先在「时间提取」页填写提取表的目标表格链接或输出文件")
        sheet = str(sheet_name or "").strip() or "提取结果"
        operation = "查询提取表"
        self._log(operation, "INFO", f"读取提取表：{path_or_url} / {sheet}")
        source = SourceConfig(
            id="extract-table",
            name="提取表",
            url=path_or_url,
            include_sheets=[sheet],
            credential_path=str(self.store.get("credential_path", "")).strip(),
        )
        extract_schema = list(self.store.get("extract_column_schema", []) or [])
        use_extract_schema = bool(self.store.get("extract_column_schema_enabled", False) and schema_field_names(extract_schema))
        records = self._reader(
            operation,
            use_schema=use_extract_schema,
            schema=extract_schema if use_extract_schema else None,
            global_excludes=[],
        ).read(source)
        for record in records:
            origin = str(record.values.get("来源", "")).strip()
            if origin:
                record.sheet_name = origin
        self._log(operation, "INFO", f"提取表读取完成：{len(records)} 行")
        return records

    @staticmethod
    def query_source_mode(source: str | None, direct: bool) -> str:
        mode = str(source or "").strip()
        if mode in {"aggregate", "extract", "direct"}:
            return mode
        return "direct" if direct else "aggregate"

    @staticmethod
    def _unique_headers(items: list[str]) -> list[str]:
        seen: set[str] = set()
        result: list[str] = []
        for item in items:
            name = str(item).strip()
            key = name.casefold()
            if name and key not in seen:
                seen.add(key)
                result.append(name)
        return result

    def _canonical_field(self, field: str) -> str:
        folded = str(field).strip().casefold()
        if not folded:
            return field
        for canonical, names in self.store.get("field_aliases", {}).items():
            options = [canonical, *names]
            if folded in {str(name).strip().casefold() for name in options}:
                return str(canonical)
        return str(field).strip()

    def _field_kind(self, field: str) -> str:
        canonical = self._canonical_field(field)
        if canonical in self.SEMANTIC_FIELD_HINTS:
            return canonical
        folded = str(field).strip().casefold()
        matches = [
            standard
            for standard, hints in self.SEMANTIC_FIELD_HINTS.items()
            if any(hint in folded for hint in hints)
        ]
        return matches[0] if len(matches) == 1 else canonical

    def _field_value(self, record: Record, field: str) -> str:
        folded = str(field).strip().casefold()
        aliases = self.store.get("field_aliases", {})
        candidates = [str(field).strip()]
        for canonical, names in aliases.items():
            options = [str(canonical), *[str(name) for name in names]]
            if folded in {name.strip().casefold() for name in options if name.strip()}:
                candidates.extend(options)
                break
        seen: set[str] = set()
        for candidate in candidates:
            key = candidate.strip().casefold()
            if not key or key in seen:
                continue
            seen.add(key)
            direct = record.values.get(candidate, "")
            if direct:
                return str(direct)
            for stored, value in record.values.items():
                if str(stored).strip().casefold() == key and str(value).strip():
                    return str(value)
        canonical = self._field_kind(field)
        hints = self.SEMANTIC_FIELD_HINTS.get(canonical, ())
        semantic_matches = [
            value
            for stored, value in record.values.items()
            if str(value).strip() and any(hint in str(stored).strip().casefold() for hint in hints)
        ]
        # Only infer a custom field when it is unambiguous. This supports names
        # such as “交教会日期” and “线索电话号码” without guessing between two dates.
        if len(semantic_matches) == 1:
            return str(semantic_matches[0])
        return ""

    def suggest_field(self, fields: list[str], preferred: str) -> str:
        """Find the best visible field for a generic/saved field name."""
        clean_fields = self._unique_headers(fields)
        folded = str(preferred).strip().casefold()
        for field in clean_fields:
            if field.casefold() == folded:
                return field
        canonical = self._field_kind(preferred)
        alias_matches = [field for field in clean_fields if self._field_kind(field) == canonical]
        if len(alias_matches) == 1:
            return alias_matches[0]
        hints = self.SEMANTIC_FIELD_HINTS.get(canonical, ())
        semantic = [field for field in clean_fields if any(hint in field.casefold() for hint in hints)]
        return semantic[0] if semantic else ""

    @staticmethod
    def corrected_source(source: str) -> str:
        if "-" not in source:
            return source
        prefix, remainder = source.split("-", 1)
        return f"{remainder}-{prefix}-专页后台"

    def _query_values(self, record: Record, field: str) -> list[str]:
        values = [self._field_value(record, field)]
        if self._field_kind(field).casefold() == "来源":
            values.extend((record.sheet_name, self.corrected_source(record.sheet_name)))
        return self._unique_headers(values)

    def query_display_value(self, record: Record, field: str) -> str:
        folded = str(field).strip().casefold()
        if folded == "修正格式":
            return self.corrected_source(record.sheet_name)
        if self._field_kind(field).casefold() == "来源":
            return record.sheet_name
        return self._field_value(record, field)

    def _peek_headers(self, target: str, sheet_name: str = "") -> list[str]:
        path = Path(str(target or "").strip())
        if not path.exists() or path.suffix.lower() not in {".xlsx", ".xlsm"}:
            return []
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            names = workbook.sheetnames
            if not names:
                return []
            title = sheet_name.strip() if sheet_name.strip() in names else names[0]
            row = next(workbook[title].iter_rows(values_only=True), None)
            return [str(cell).strip() for cell in (row or []) if str(cell or "").strip()]
        finally:
            workbook.close()

    def list_query_fields(
        self,
        mode: str = "extract",
        extract_target: str = "",
        extract_sheet: str = "",
        source_id: str = "",
    ) -> list[str]:
        fallback = list(self.store.get("field_aliases", {}).keys())
        mode = self.query_source_mode(mode, False)
        if mode == "extract":
            schema_names = schema_field_names(self.store.get("extract_column_schema", []))
            if self.store.get("extract_column_schema_enabled", False) and schema_names:
                return self._unique_headers([*schema_names, *fallback])
            peeked = self._peek_headers(extract_target, extract_sheet)
            if peeked:
                return self._unique_headers(peeked)
            return self._unique_headers(["来源", *fallback])
        if mode == "direct":
            sources = [source for source in self.store.load_sources() if source.enabled]
            if source_id.strip():
                sources = [source for source in sources if source.id == source_id.strip()]
            headers: list[str] = []
            for source in sources:
                if source.column_schema_enabled and source.column_schema:
                    headers.extend(schema_field_names(source.column_schema))
                    continue
                peeked = self._peek_headers(source.url, (source.include_sheets or [""])[0])
                if peeked:
                    headers.extend(peeked)
                elif self.store.get("column_schema_enabled", False):
                    headers.extend(schema_field_names(self.store.get("column_schema", [])))
            return self._unique_headers(headers) or self._unique_headers(fallback)
        headers = []
        for record in self.database.sample_records(20):
            headers.extend(record.values.keys())
        if self.store.get("column_schema_enabled", False):
            headers.extend(schema_field_names(self.store.get("column_schema", [])))
        return self._unique_headers(headers) or self._unique_headers(fallback)

    def query(
        self,
        field: str,
        value: str,
        direct: bool = False,
        exact: bool = True,
        source: str | None = None,
        extract_target: str = "",
        extract_sheet: str = "",
        source_id: str = "",
        date_field: str = "",
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[Record]:
        results = self.query_many(
            field, [value], direct, exact, source, extract_target, extract_sheet, source_id,
            date_field, start_date, end_date,
        )
        return [record for _, record in results if record is not None]

    def query_many(
        self,
        field: str,
        values: list[str],
        direct: bool = False,
        exact: bool = True,
        source: str | None = None,
        extract_target: str = "",
        extract_sheet: str = "",
        source_id: str = "",
        date_field: str = "",
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[tuple[str, Record | None]]:
        mode = self.query_source_mode(source, direct)
        operation = {"direct": "批量直接查询", "extract": "批量提取表查询"}.get(mode, "批量汇总库查询")
        queries = list(dict.fromkeys(value.strip() for value in values if value.strip()))
        if not queries:
            raise ValueError("没有可查询的号码")
        self._log(operation, "INFO", f"开始批量查询：{len(queries)} 个值，字段“{field}”")
        if mode == "extract":
            records = self.read_extract_table(extract_target, extract_sheet)
        elif mode == "direct":
            records = self.read_sources(operation, source_id)
        else:
            records = self.database.all_records()
        if date_field.strip():
            if start_date is None or end_date is None:
                raise ValueError("日期范围不完整")
            if start_date > end_date:
                raise ValueError("开始日期不能晚于结束日期")
            before = len(records)
            filtered: list[Record] = []
            valid_date_count = 0
            for record in records:
                current = parse_date(self._field_value(record, date_field))
                if current is not None:
                    valid_date_count += 1
                if current is not None and start_date <= current <= end_date:
                    filtered.append(record)
            records = filtered
            self._log(
                operation,
                "INFO",
                f"日期限制：字段“{date_field}”，{start_date.isoformat()} 至 {end_date.isoformat()}，保留 {len(records)}/{before} 行",
            )
            if before and valid_date_count == 0:
                raise ValueError(
                    f"日期字段“{date_field}”在读取的 {before} 行中没有可识别日期。"
                    "请从下拉框选择当前数据表实际使用的日期列。"
                )
        canonical = self._field_kind(field)
        results: list[tuple[str, Record | None]] = []
        matched_inputs = 0
        for query_value in queries:
            needle = self._match_value(canonical, query_value)
            matches = []
            for record in records:
                haystacks = [self._match_value(canonical, value) for value in self._query_values(record, field)]
                if any((haystack == needle) if exact else (needle in haystack) for haystack in haystacks):
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
        text = unicodedata.normalize("NFKC", str(value)).strip().casefold().lstrip("'")
        text = re.sub(r"[\u200b-\u200f\u2060\ufeff]", "", text)
        if field.strip().casefold() not in {"号码", "手机号", "手机号码", "电话", "联系电话", "phone", "number"}:
            return re.sub(r"\s+", "", text)
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
            current = parse_date(self._field_value(record, date_field))
            if current is None:
                skipped_invalid += 1
                continue
            if not start <= current <= end:
                continue
            parts = [self._field_value(record, field) for field in dedup_fields]
            if not any(parts):
                parts = [record.spreadsheet_id, record.sheet_name, str(record.row_number), record.row_hash]
            key = hashlib.sha256(json.dumps(parts, ensure_ascii=False).encode("utf-8")).hexdigest()
            if self.store.was_extracted(key):
                skipped_duplicate += 1
                continue
            selected.append((record, key))
        if records and skipped_invalid == len(records):
            raise ValueError(
                f"日期字段“{date_field}”在读取的 {len(records)} 行中没有可识别日期。"
                "请在“日期字段”中选择数据源实际使用的日期列。"
            )
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
        if self.store.get("extract_column_schema_enabled", False):
            configured = schema_field_names(self.store.get("extract_column_schema", []))
            if configured:
                return ["来源", *[name for name in configured if name != "来源"]]
        if self.store.get("column_schema_enabled", False):
            configured = schema_field_names(self.store.get("column_schema", []))
            if configured:
                return ["来源", *[name for name in configured if name != "来源"]]
        headers, _ = self._export_rows(records)
        return headers

    @classmethod
    def _canonical_header(cls, header: str, aliases: dict[str, list[str]]) -> str:
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
        semantic = [
            standard
            for standard, hints in cls.SEMANTIC_FIELD_HINTS.items()
            if any(hint in folded for hint in hints)
        ]
        if len(semantic) == 1:
            return semantic[0]
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
                    value = record.values.get(canonical, record.values.get(header, ""))
                    if not str(value).strip() and canonical in cls.SEMANTIC_FIELD_HINTS:
                        matches = [
                            candidate
                            for stored, candidate in record.values.items()
                            if str(candidate).strip()
                            and any(hint in str(stored).strip().casefold() for hint in cls.SEMANTIC_FIELD_HINTS[canonical])
                        ]
                        if len(matches) == 1:
                            value = matches[0]
                    row.append(value)
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
