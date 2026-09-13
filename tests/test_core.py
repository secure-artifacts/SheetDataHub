from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import openpyxl

from sheet_hub.config_store import ConfigStore
from sheet_hub.database import AggregateDatabase
from sheet_hub.engine import DataEngine, parse_date
from sheet_hub.models import Record, SourceConfig
from sheet_hub.source_reader import SourceReader, canonicalize, choose_sheets, google_retry


class RuleTests(unittest.TestCase):
    def test_empty_include_runs_all_except_excluded(self):
        selected, missing = choose_sheets(["订单", "统计", "客户"], [], ["统计"])
        self.assertEqual(selected, ["订单", "客户"])
        self.assertEqual(missing, [])

    def test_include_only_and_exclude_wins(self):
        selected, missing = choose_sheets(["订单A", "订单B"], ["订单A", "订单B", "不存在"], ["订单B"])
        self.assertEqual(selected, ["订单A"])
        self.assertEqual(missing, ["不存在"])

    def test_space_separated_sheet_exclusions_are_supported(self):
        selected, _ = choose_sheets(["数据", "Index", "清理", "备份链接"], [], ["Index 清理 备份链接"])
        self.assertEqual(selected, ["数据"])

    def test_alias_mapping(self):
        result = canonicalize({"手机号": 13800138000, "渠道": "广告"}, {"号码": ["手机号"], "来源": ["渠道"]})
        self.assertEqual(result["号码"], "13800138000")
        self.assertEqual(result["来源"], "广告")

    def test_date_parsing(self):
        self.assertEqual(parse_date("2026/09/13"), date(2026, 9, 13))
        self.assertIsNone(parse_date("not-a-date"))

    def test_google_429_is_retried(self):
        class QuotaError(Exception):
            response = SimpleNamespace(status_code=429)

        calls = {"count": 0}

        def operation():
            calls["count"] += 1
            if calls["count"] == 1:
                raise QuotaError()
            return "ok"

        with patch("sheet_hub.source_reader.time.sleep"), patch("sheet_hub.source_reader.random.random", return_value=0):
            self.assertEqual(google_retry(operation), "ok")
        self.assertEqual(calls["count"], 2)

    def test_phone_matching_normalizes_common_formats(self):
        self.assertEqual(DataEngine._match_value("号码", "258-851-758692"), "258851758692")
        self.assertEqual(DataEngine._match_value("号码", "258851758692.0"), "258851758692")
        self.assertEqual(DataEngine._match_value("号码", "2.58851758692E+11"), "258851758692")

    def test_existing_aliases_are_migrated(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ConfigStore(directory)
            aliases = store.get("field_aliases")
            aliases["号码"] = ["号码", "手机号"]
            store.set("field_aliases", aliases)
            reopened = ConfigStore(directory)
            self.assertIn("手机号码", reopened.get("field_aliases")["号码"])


class DatabaseTests(unittest.TestCase):
    def test_sharding_and_query(self):
        with tempfile.TemporaryDirectory() as directory:
            db = AggregateDatabase(directory, 1000)
            records = [Record("s", "源", "g", "订单", index, {"号码": str(index)}, str(index)) for index in range(1001)]
            rows, databases = db.replace_all(records)
            self.assertEqual(rows, 1001)
            self.assertEqual(databases, 2)
            self.assertEqual(db.query("号码", "1000")[0].row_number, 1000)

    def test_extract_dedup(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            store = ConfigStore(root / "config")
            engine = DataEngine(store)
            records = [Record("s", "源", "g", "订单", 2, {"日期": "2026-09-13", "号码": "123"}, "hash")]
            engine.database.replace_all(records)
            output1 = root / "one.xlsx"
            output2 = root / "two.xlsx"
            first = engine.extract("日期", date(2026, 9, 1), date(2026, 9, 30), output1, ["号码", "日期"])
            second = engine.extract("日期", date(2026, 9, 1), date(2026, 9, 30), output2, ["号码", "日期"])
            self.assertEqual(first["written"], 1)
            self.assertEqual(second["written"], 0)
            self.assertTrue(output2.exists())

    def test_batch_query_keeps_missing_numbers(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ConfigStore(Path(directory) / "config")
            engine = DataEngine(store)
            engine.database.replace_all([
                Record("s", "源", "g", "1008-李薇", 2, {"号码": "123", "名字": "李薇", "日期": "2026-09-13"}, "h1")
            ])
            results = engine.query_many("号码", ["123", "999"])
            self.assertEqual(results[0][1].sheet_name, "1008-李薇")
            self.assertIsNone(results[1][1])

    def test_extract_uses_configured_sheet_and_source_first(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "result.xlsx"
            records = [Record("s", "源", "g", "渠道Sheet", 2, {"号码": "123", "日期": "2026-09-13"}, "h")]
            DataEngine._write_xlsx(path, records, "指定工作表")
            book = openpyxl.load_workbook(path, read_only=True, data_only=True)
            self.assertEqual(book.sheetnames, ["指定工作表"])
            rows = list(book["指定工作表"].iter_rows(values_only=True))
            self.assertEqual(rows[0][0], "来源")
            self.assertEqual(rows[1][0], "渠道Sheet")
            book.close()

    def test_export_header_aliases_and_order_are_compatible(self):
        aliases = {
            "名字": ["姓名"],
            "号码": ["手机号码"],
            "日期": ["日期"],
        }
        existing = ["来源", "专页ID", "姓名", "手机号码", "日期"]
        required = ["来源", "日期", "名字", "号码", "专页ID"]
        self.assertTrue(DataEngine._headers_compatible(existing, required, aliases))
        records = [Record(
            "s", "源", "g", "1008-李薇", 2,
            {"专页ID": "42", "名字": "李薇", "号码": "258851758692", "日期": "2026-09-13"},
            "h",
        )]
        self.assertEqual(
            DataEngine._rows_for_headers(records, existing, aliases)[0],
            ["1008-李薇", "42", "李薇", "258851758692", "2026-09-13"],
        )


class WorkbookTests(unittest.TestCase):
    def test_local_workbook_source(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "sample.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "订单"
            sheet.append(["日期", "手机号", "渠道"])
            sheet.append(["2026-09-13", "10086", "搜索"])
            stats = book.create_sheet("统计")
            stats.append(["日期", "手机号"])
            stats.append(["2026-09-13", "should_skip"])
            book.save(path)
            reader = SourceReader({"号码": ["手机号"], "来源": ["渠道"], "日期": ["日期"]}, ["统计"])
            source = SourceConfig("id", "测试", str(path))
            records = reader.read(source)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0].values["号码"], "10086")

    def test_manual_column_count_and_headers(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "manual.xlsx"
            book = openpyxl.Workbook()
            sheet = book.active
            sheet.title = "数据"
            sheet.append(["原表头A", "原表头B", "不读取的列"])
            sheet.append(["page-1", "13800138000", "ignore-me"])
            book.save(path)
            reader = SourceReader(
                {"号码": ["手机号码"]},
                [],
                column_schema=["专页ID", "手机号码"],
            )
            records = reader.read(SourceConfig("id", "手动列", str(path)))
            self.assertEqual(records[0].values["专页ID"], "page-1")
            self.assertEqual(records[0].values["号码"], "13800138000")
            self.assertNotIn("不读取的列", records[0].values)

    def test_private_source_uses_one_batch_request(self):
        class FakeSpreadsheet:
            def __init__(self):
                self.batch_calls = 0

            def fetch_sheet_metadata(self, key, params=None):
                return {"sheets": [
                    {"properties": {"title": "A"}},
                    {"properties": {"title": "B"}},
                ]}

            def values_batch_get(self, key, ranges, params=None):
                self.batch_calls += 1
                return {"valueRanges": [
                    {"values": [["手机号"], ["111"]]},
                    {"values": [["手机号"], ["222"]]},
                ]}

        spreadsheet = FakeSpreadsheet()
        client = SimpleNamespace(http_client=spreadsheet)
        source = SourceConfig("id", "私有表", "https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz", credential_path="fake.json")
        reader = SourceReader({"号码": ["手机号"]}, [])
        with patch.object(SourceReader, "_gspread_client", return_value=client):
            records = reader.read(source)
        self.assertEqual(spreadsheet.batch_calls, 1)
        self.assertEqual([record.values["号码"] for record in records], ["111", "222"])

    def test_google_output_appends_header_and_rows(self):
        class FakeHttp:
            def __init__(self):
                self.appended = []

            def fetch_sheet_metadata(self, key, params=None):
                return {"sheets": [{"properties": {"sheetId": 1, "title": "目标页"}}]}

            def values_get(self, key, range_name, params=None):
                return {"values": []}

            def values_append(self, key, range_name, params=None, body=None):
                self.appended.extend(body["values"])

        http = FakeHttp()
        client = SimpleNamespace(http_client=http)
        with tempfile.TemporaryDirectory() as directory:
            store = ConfigStore(Path(directory) / "config")
            store.set("credential_path", "fake.json")
            engine = DataEngine(store)
            records = [Record("s", "源", "g", "渠道A", 2, {"号码": "123"}, "h")]
            with patch.object(SourceReader, "_gspread_client", return_value=client):
                engine._write_google_sheet(
                    "https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz",
                    "目标页",
                    records,
                    "测试",
                )
        self.assertEqual(http.appended[0], ["来源", "号码"])
        self.assertEqual(http.appended[1], ["渠道A", "123"])

    def test_google_output_uses_existing_alias_header_order(self):
        existing = ["来源", "专页ID", "姓名", "标签", "订阅时间", "性别", "评论贴文", "手机号码", "日期"]

        class FakeHttp:
            def __init__(self):
                self.appended = []

            def fetch_sheet_metadata(self, key, params=None):
                return {"sheets": [{"properties": {"sheetId": 1, "title": "测试"}}]}

            def values_get(self, key, range_name, params=None):
                return {"values": [existing]}

            def values_append(self, key, range_name, params=None, body=None):
                self.appended.extend(body["values"])

        http = FakeHttp()
        client = SimpleNamespace(http_client=http)
        with tempfile.TemporaryDirectory() as directory:
            store = ConfigStore(Path(directory) / "config")
            store.set("credential_path", "fake.json")
            store.set("column_schema_enabled", True)
            store.set("column_schema", existing[1:])
            engine = DataEngine(store)
            records = [Record(
                "s", "源", "g", "1008-李薇", 2,
                {
                    "专页ID": "42", "名字": "李薇", "标签": "A", "订阅时间": "12:00",
                    "性别": "女", "评论贴文": "https://example.test", "号码": "258851758692",
                    "日期": "2026-09-13",
                },
                "h",
            )]
            with patch.object(SourceReader, "_gspread_client", return_value=client):
                engine._write_google_sheet(
                    "https://docs.google.com/spreadsheets/d/abcdefghijklmnopqrstuvwxyz",
                    "测试",
                    records,
                    "测试",
                )
        self.assertEqual(http.appended[0], [
            "1008-李薇", "42", "李薇", "A", "12:00", "女", "https://example.test",
            "258851758692", "2026-09-13",
        ])


if __name__ == "__main__":
    unittest.main()
