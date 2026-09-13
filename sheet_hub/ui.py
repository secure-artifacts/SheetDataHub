from __future__ import annotations

import json
import sys
import uuid
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, QThread, Qt, Signal
from PySide6.QtGui import QColor, QFontDatabase, QIcon, QPixmap
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .config_store import ConfigStore
from .engine import DataEngine
from .models import Record, SourceConfig
from .source_reader import SourceReader, split_names


APP_TITLE = "表数通"


class TaskThread(QThread):
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, job: Callable[[], object], parent=None):
        super().__init__(parent)
        self.job = job

    def run(self) -> None:
        try:
            self.succeeded.emit(self.job())
        except Exception as exc:
            self.failed.emit(str(exc))


class SourceDialog(QDialog):
    def __init__(self, store: ConfigStore, source: SourceConfig | None = None, parent=None):
        super().__init__(parent)
        self.store = store
        self.source = source
        self.setWindowTitle("编辑数据源" if source else "添加数据源")
        self.setMinimumWidth(650)
        form = QFormLayout(self)
        self.name = QLineEdit(source.name if source else "")
        self.url = QLineEdit(source.url if source else "")
        self.url.setPlaceholderText("Google 表格链接；也支持本地 .xlsx 文件")
        self.include = QTextEdit("\n".join(source.include_sheets) if source else "")
        self.include.setMaximumHeight(85)
        self.include.setPlaceholderText("留空＝遍历全部；多个名称每行一个")
        self.exclude = QTextEdit("\n".join(source.exclude_sheets) if source else "")
        self.exclude.setMaximumHeight(85)
        self.exclude.setPlaceholderText("多个名称每行一个；排除规则优先")
        self.header_row = QSpinBox()
        self.header_row.setRange(1, 100)
        self.header_row.setValue(source.header_row if source else 1)
        self.credential = QLineEdit(source.credential_path if source else "")
        credential_button = QPushButton("选择…")
        credential_button.clicked.connect(self.choose_credential)
        credential_row = QHBoxLayout()
        credential_row.addWidget(self.credential, 1)
        credential_row.addWidget(credential_button)
        self.enabled = QCheckBox("启用此数据源")
        self.enabled.setChecked(source.enabled if source else True)
        form.addRow("数据源名称*", self.name)
        form.addRow("表格链接/文件*", self.url)
        form.addRow("指定 Sheet", self.include)
        form.addRow("排除 Sheet", self.exclude)
        form.addRow("表头所在行", self.header_row)
        form.addRow("服务账号 JSON", credential_row)
        form.addRow("", self.enabled)
        hint = QLabel("私有 Google 表格需要服务账号；留空时按公开表格读取。")
        hint.setObjectName("muted")
        form.addRow("", hint)
        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        form.addRow(buttons)

    def choose_credential(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择服务账号 JSON", "", "JSON 文件 (*.json)")
        if path:
            self.credential.setText(path)

    def validate_and_accept(self) -> None:
        if not self.name.text().strip() or not self.url.text().strip():
            QMessageBox.warning(self, "缺少信息", "请填写数据源名称和表格链接。")
            return
        self.accept()

    def value(self) -> SourceConfig:
        return SourceConfig(
            id=self.source.id if self.source else uuid.uuid4().hex,
            name=self.name.text().strip(),
            url=self.url.text().strip(),
            include_sheets=split_names(self.include.toPlainText()),
            exclude_sheets=split_names(self.exclude.toPlainText()),
            header_row=self.header_row.value(),
            enabled=self.enabled.isChecked(),
            credential_path=self.credential.text().strip(),
        )


class MainWindow(QMainWindow):
    def __init__(self, store: ConfigStore, icon_path: Path | None = None):
        super().__init__()
        self.store = store
        self.tasks: list[TaskThread] = []
        self.icon_path = icon_path
        self.setWindowTitle(f"{APP_TITLE} · 多表数据查询与提取")
        self.resize(1220, 780)
        self.setMinimumSize(980, 650)
        if icon_path and icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self._build_ui()
        self.refresh_sources()
        self.refresh_field_controls()
        self.refresh_logs()

    def _build_ui(self) -> None:
        central = QWidget()
        outer = QHBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        sidebar = QFrame()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(210)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(20, 24, 20, 20)
        brand_row = QHBoxLayout()
        if self.icon_path and self.icon_path.exists():
            logo = QLabel()
            logo.setPixmap(QPixmap(str(self.icon_path)).scaled(46, 46, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            brand_row.addWidget(logo)
        brand_text = QVBoxLayout()
        brand = QLabel(APP_TITLE)
        brand.setObjectName("brand")
        subtitle = QLabel("多表数据工作台")
        subtitle.setObjectName("sidebarMuted")
        brand_text.addWidget(brand)
        brand_text.addWidget(subtitle)
        brand_row.addLayout(brand_text, 1)
        self.nav = QListWidget()
        self.nav.setObjectName("navigation")
        self.nav.addItems(["数据源", "汇总同步", "数据查询", "时间提取", "字段配置", "运行日志", "设置"])
        self.nav.setCurrentRow(0)
        sidebar_layout.addLayout(brand_row)
        sidebar_layout.addSpacing(22)
        sidebar_layout.addWidget(self.nav, 1)
        version = QLabel("v1.1.2")
        version.setObjectName("sidebarMuted")
        sidebar_layout.addWidget(version)
        self.pages = QStackedWidget()
        self.pages.addWidget(self._sources_page())
        self.pages.addWidget(self._sync_page())
        self.pages.addWidget(self._query_page())
        self.pages.addWidget(self._extract_page())
        self.pages.addWidget(self._fields_page())
        self.pages.addWidget(self._logs_page())
        self.pages.addWidget(self._settings_page())
        self.nav.currentRowChanged.connect(self.pages.setCurrentIndex)
        outer.addWidget(sidebar)
        outer.addWidget(self.pages, 1)
        self.setCentralWidget(central)
        self.statusBar().showMessage("就绪")

    def _page(self, title: str, description: str) -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(30, 26, 30, 24)
        title_label = QLabel(title)
        title_label.setObjectName("pageTitle")
        desc_label = QLabel(description)
        desc_label.setObjectName("muted")
        desc_label.setWordWrap(True)
        layout.addWidget(title_label)
        layout.addWidget(desc_label)
        layout.addSpacing(14)
        return page, layout

    def _sources_page(self) -> QWidget:
        page, layout = self._page("数据源", "配置一个或多个 Google 表格链接；每个链接可以指定或排除子 Sheet。")
        actions = QHBoxLayout()
        add = QPushButton("＋ 添加数据源")
        add.setObjectName("primary")
        edit = QPushButton("编辑")
        remove = QPushButton("删除")
        scan = QPushButton("扫描 Sheet")
        add.clicked.connect(self.add_source)
        edit.clicked.connect(self.edit_source)
        remove.clicked.connect(self.delete_source)
        scan.clicked.connect(self.scan_source)
        actions.addWidget(add)
        actions.addWidget(edit)
        actions.addWidget(remove)
        actions.addWidget(scan)
        actions.addStretch()
        self.source_table = QTableWidget(0, 7)
        self.source_table.setHorizontalHeaderLabels(["启用", "名称", "表格链接", "指定 Sheet", "排除 Sheet", "表头行", "认证"])
        self.source_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.source_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.source_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.source_table.verticalHeader().setVisible(False)
        self.source_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.source_table.doubleClicked.connect(self.edit_source)
        rule = QLabel("读取规则：指定名称留空时遍历全部；填写后仅抓取指定项；排除项始终优先。")
        rule.setObjectName("infoBox")
        layout.addLayout(actions)
        layout.addWidget(self.source_table, 1)
        layout.addWidget(rule)
        return page

    def _sync_page(self) -> QWidget:
        page, layout = self._page("汇总同步", "读取所有已启用数据源。是否写入本地汇总库由下方开关决定。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.write_aggregate = QCheckBox("读取完成后写入汇总数据库")
        self.write_aggregate.setChecked(bool(self.store.get("write_aggregate", True)))
        self.write_aggregate.toggled.connect(lambda checked: self.store.set("write_aggregate", checked))
        self.max_rows = QSpinBox()
        self.max_rows.setRange(1000, 10_000_000)
        self.max_rows.setSingleStep(10000)
        self.max_rows.setValue(int(self.store.get("max_rows_per_db", 500000)))
        self.max_rows.valueChanged.connect(lambda value: self.store.set("max_rows_per_db", value))
        form.addRow("汇总方式", self.write_aggregate)
        form.addRow("单库最大行数", self.max_rows)
        hint = QLabel("超过单库上限时会自动生成 aggregate_002、aggregate_003…，查询时自动跨库。")
        hint.setObjectName("muted")
        form.addRow("", hint)
        self.sync_button = QPushButton("开始读取与同步")
        self.sync_button.setObjectName("primary")
        self.sync_button.clicked.connect(self.run_sync)
        self.sync_status = QTextEdit()
        self.sync_status.setReadOnly(True)
        self.sync_status.setPlaceholderText("运行明细会显示在这里，并同步写入运行日志。")
        layout.addWidget(card)
        layout.addWidget(self.sync_button, 0, Qt.AlignLeft)
        layout.addWidget(self.sync_status, 1)
        return page

    def _query_page(self) -> QWidget:
        page, layout = self._page("数据查询", "按任意标准字段查询；可选择汇总库查询或直接遍历数据源。")
        bar = QHBoxLayout()
        self.query_field = QComboBox()
        self.query_value = QTextEdit()
        self.query_value.setMaximumHeight(78)
        self.query_value.setPlaceholderText("每行一个号码，也支持逗号分隔批量粘贴")
        self.query_mode = QComboBox()
        self.query_mode.addItems(["查询汇总数据库", "直接查询数据源"])
        self.query_exact = QCheckBox("精确匹配")
        self.query_exact.setChecked(True)
        button = QPushButton("查询")
        button.setObjectName("primary")
        button.clicked.connect(self.run_query)
        self.copy_query_button = QPushButton("一键复制号码和修正格式")
        self.copy_query_button.setEnabled(False)
        self.copy_query_button.clicked.connect(self.copy_query_results)
        bar.addWidget(QLabel("查询字段"))
        bar.addWidget(self.query_field)
        bar.addWidget(self.query_value, 1)
        bar.addWidget(self.query_mode)
        bar.addWidget(self.query_exact)
        bar.addWidget(button)
        bar.addWidget(self.copy_query_button)
        self.query_table = QTableWidget()
        self.query_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.query_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.query_table.verticalHeader().setVisible(False)
        self.query_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        layout.addLayout(bar)
        layout.addWidget(self.query_table, 1)
        return page

    def _extract_page(self) -> QWidget:
        page, layout = self._page("时间提取", "按日期范围提取数据并写入独立 Excel；历史提取记录会自动排重。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.extract_date_field = QComboBox()
        self.start_date = QDateEdit(QDate.currentDate().addMonths(-1))
        self.end_date = QDateEdit(QDate.currentDate())
        for widget in (self.start_date, self.end_date):
            widget.setCalendarPopup(True)
            widget.setDisplayFormat("yyyy-MM-dd")
        date_row = QHBoxLayout()
        date_row.addWidget(self.start_date)
        date_row.addWidget(QLabel("至"))
        date_row.addWidget(self.end_date)
        date_row.addStretch()
        self.dedup_fields = QLineEdit("号码,日期")
        self.dedup_fields.setPlaceholderText("多个字段用逗号分隔；留空时使用源位置和行指纹")
        self.extract_mode = QComboBox()
        self.extract_mode.addItems(["从汇总数据库提取", "直接从数据源提取"])
        self.output_type = QComboBox()
        self.output_type.addItems(["本地 Excel 文件", "Google 表格链接"])
        self.output_sheet_name = QLineEdit(str(self.store.get("google_output_sheet", "提取结果")))
        self.output_sheet_name.setPlaceholderText("例如：查询结果")
        self.google_output_url = QLineEdit(str(self.store.get("google_output_url", "")))
        self.google_output_url.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        self.output_path = QLineEdit(str(Path.home() / "Desktop" / "提取结果.xlsx"))
        choose = QPushButton("选择…")
        choose.clicked.connect(self.choose_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_path, 1)
        output_row.addWidget(choose)
        form.addRow("日期字段", self.extract_date_field)
        form.addRow("日期范围", date_row)
        form.addRow("去重字段", self.dedup_fields)
        form.addRow("数据来源", self.extract_mode)
        form.addRow("输出目标", self.output_type)
        form.addRow("目标表格链接", self.google_output_url)
        form.addRow("目标工作表", self.output_sheet_name)
        form.addRow("输出文件", output_row)
        self.output_type.currentIndexChanged.connect(self.update_output_destination)
        self.local_output_widgets = [self.output_path, choose]
        self.update_output_destination()
        button = QPushButton("开始提取")
        button.setObjectName("primary")
        button.clicked.connect(self.run_extract)
        self.extract_status = QLabel("等待运行")
        self.extract_status.setObjectName("muted")
        layout.addWidget(card)
        layout.addWidget(button, 0, Qt.AlignLeft)
        layout.addWidget(self.extract_status)
        layout.addStretch()
        return page

    def _fields_page(self) -> QWidget:
        page, layout = self._page("列与字段配置", "可以明确指定读取列数和每一列的表头；不启用时自动使用数据源原始表头。")
        schema_card = QFrame()
        schema_card.setObjectName("card")
        schema_layout = QVBoxLayout(schema_card)
        schema_actions = QHBoxLayout()
        self.schema_enabled = QCheckBox("启用手动列结构")
        self.schema_enabled.setChecked(bool(self.store.get("column_schema_enabled", False)))
        schema = self.store.get("column_schema", [])
        self.schema_count = QSpinBox()
        self.schema_count.setRange(1, 200)
        self.schema_count.setValue(len(schema) or 8)
        example = QPushButton("填入示例 8 列")
        example.clicked.connect(self.fill_example_schema)
        schema_actions.addWidget(self.schema_enabled)
        schema_actions.addSpacing(18)
        schema_actions.addWidget(QLabel("读取列数"))
        schema_actions.addWidget(self.schema_count)
        schema_actions.addWidget(example)
        schema_actions.addStretch()
        self.schema_table = QTableWidget(0, 2)
        self.schema_table.setHorizontalHeaderLabels(["列", "指定表头名称"])
        self.schema_table.setMaximumHeight(245)
        self.schema_table.verticalHeader().setVisible(False)
        self.schema_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.schema_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        schema_layout.addLayout(schema_actions)
        schema_layout.addWidget(self.schema_table)
        schema_hint = QLabel("启用后只读取前 N 列，并按这里的名称识别 A、B、C…列；名称留空时保留原表头。")
        schema_hint.setObjectName("muted")
        schema_layout.addWidget(schema_hint)
        layout.addWidget(schema_card)
        layout.addWidget(QLabel("标准字段与表头别名"))
        self.fields_table = QTableWidget(0, 2)
        self.fields_table.setHorizontalHeaderLabels(["标准字段", "可匹配的表头别名（逗号分隔）"])
        self.fields_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        actions = QHBoxLayout()
        add = QPushButton("添加字段")
        remove = QPushButton("删除选中")
        save = QPushButton("保存列与字段配置")
        save.setObjectName("primary")
        add.clicked.connect(self.add_field_row)
        remove.clicked.connect(self.remove_field_row)
        save.clicked.connect(self.save_fields)
        actions.addWidget(add)
        actions.addWidget(remove)
        actions.addStretch()
        actions.addWidget(save)
        layout.addWidget(self.fields_table, 1)
        layout.addLayout(actions)
        self.schema_count.valueChanged.connect(self.resize_schema_table)
        self.schema_enabled.toggled.connect(self.update_schema_enabled)
        self.load_fields_table()
        return page

    def _logs_page(self) -> QWidget:
        page, layout = self._page("运行日志", "查看读取、匹配、写入、去重和错误的详细记录。")
        actions = QHBoxLayout()
        refresh = QPushButton("刷新")
        clear = QPushButton("清空日志")
        refresh.clicked.connect(self.refresh_logs)
        clear.clicked.connect(self.clear_logs)
        actions.addWidget(refresh)
        actions.addWidget(clear)
        actions.addStretch()
        self.log_table = QTableWidget(0, 5)
        self.log_table.setHorizontalHeaderLabels(["时间", "级别", "操作", "消息", "详情"])
        self.log_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.log_table.verticalHeader().setVisible(False)
        self.log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        layout.addLayout(actions)
        layout.addWidget(self.log_table, 1)
        return page

    def _settings_page(self) -> QWidget:
        page, layout = self._page("设置", "配置所有数据源共用的默认排除项和认证文件。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.global_excludes = QTextEdit("\n".join(self.store.get("global_excludes", [])))
        self.global_excludes.setMaximumHeight(130)
        self.default_credential = QLineEdit(self.store.get("credential_path", ""))
        choose = QPushButton("选择…")
        choose.clicked.connect(self.choose_default_credential)
        credential_row = QHBoxLayout()
        credential_row.addWidget(self.default_credential, 1)
        credential_row.addWidget(choose)
        form.addRow("全局排除 Sheet", self.global_excludes)
        form.addRow("默认服务账号 JSON", credential_row)
        save = QPushButton("保存设置")
        save.setObjectName("primary")
        save.clicked.connect(self.save_settings)
        data_path = QLabel(str(self.store.data_dir))
        data_path.setTextInteractionFlags(Qt.TextSelectableByMouse)
        form.addRow("本地数据目录", data_path)
        layout.addWidget(card)
        layout.addWidget(save, 0, Qt.AlignLeft)
        layout.addStretch()
        return page

    def selected_source(self) -> SourceConfig | None:
        row = self.source_table.currentRow()
        sources = self.store.load_sources()
        return sources[row] if 0 <= row < len(sources) else None

    def refresh_sources(self) -> None:
        sources = self.store.load_sources()
        self.source_table.setRowCount(len(sources))
        for row, source in enumerate(sources):
            values = [
                "是" if source.enabled else "否", source.name, source.url,
                "、".join(source.include_sheets) or "全部", "、".join(source.exclude_sheets) or "—",
                str(source.header_row), "服务账号" if source.credential_path else "公开读取",
            ]
            for column, value in enumerate(values):
                self.source_table.setItem(row, column, QTableWidgetItem(value))

    def add_source(self) -> None:
        dialog = SourceDialog(self.store, parent=self)
        if dialog.exec():
            value = dialog.value()
            if not value.credential_path:
                value.credential_path = self.store.get("credential_path", "")
            self.store.save_source(value)
            self.refresh_sources()

    def edit_source(self) -> None:
        source = self.selected_source()
        if not source:
            QMessageBox.information(self, "请选择", "请先选择一个数据源。")
            return
        dialog = SourceDialog(self.store, source, self)
        if dialog.exec():
            self.store.save_source(dialog.value())
            self.refresh_sources()

    def delete_source(self) -> None:
        source = self.selected_source()
        if not source:
            return
        if QMessageBox.question(self, "确认删除", f"确定删除数据源“{source.name}”吗？") == QMessageBox.Yes:
            self.store.delete_source(source.id)
            self.refresh_sources()

    def scan_source(self) -> None:
        source = self.selected_source()
        if not source:
            QMessageBox.information(self, "请选择", "请先选择一个数据源。")
            return
        reader = SourceReader(self.store.get("field_aliases", {}), self.store.get("global_excludes", []))
        self.run_task(lambda: reader.list_sheets(source), self.scan_finished, "正在扫描 Sheet…")

    def scan_finished(self, names: list[str]) -> None:
        QMessageBox.information(self, "Sheet 列表", "\n".join(names) if names else "没有发现子 Sheet。")

    def run_sync(self) -> None:
        self.sync_status.clear()
        self.sync_button.setEnabled(False)
        engine = DataEngine(self.store)
        self.run_task(engine.sync, self.sync_finished, "正在读取与同步…", self.sync_button)

    def sync_finished(self, result: dict[str, int]) -> None:
        self.refresh_logs()
        recent = list(reversed(self.store.read_logs(100)))
        self.sync_status.setPlainText("\n".join(
            f"[{row['created_at']}] {row['level']} · {row['message']}" for row in recent
            if row["operation"] in {"汇总同步", "读取数据源"}
        ))
        if result["databases"]:
            QMessageBox.information(self, "同步完成", f"已写入 {result['rows']} 行，使用 {result['databases']} 个数据库文件。")
        else:
            QMessageBox.information(self, "读取完成", f"已读取 {result['rows']} 行；未写入汇总库。")

    def run_query(self) -> None:
        values = split_names(self.query_value.toPlainText())
        if not values:
            QMessageBox.warning(self, "请输入", "请输入一个或多个查询号码。")
            return
        engine = DataEngine(self.store)
        field = self.query_field.currentText()
        direct = self.query_mode.currentIndex() == 1
        exact = self.query_exact.isChecked()
        self.run_task(lambda: engine.query_many(field, values, direct, exact), self.show_query_results, "正在批量查询…")

    @staticmethod
    def corrected_source(source: str) -> str:
        if "-" not in source:
            return source
        prefix, remainder = source.split("-", 1)
        return f"{remainder}-{prefix}-专页后台"

    @staticmethod
    def record_value(record: Record, *names: str) -> str:
        for name in names:
            value = record.values.get(name, "")
            if value:
                return value
        return ""

    def show_query_results(self, results: list[tuple[str, Record | None]]) -> None:
        fields = ["输入电话号码", "专页ID", "姓名", "评论贴文", "号码", "日期", "修正格式"]
        self.query_table.setColumnCount(len(fields))
        self.query_table.setHorizontalHeaderLabels(fields)
        self.query_table.setRowCount(len(results))
        self.query_copy_rows: list[tuple[str, str]] = []
        found = 0
        for row, (query_value, record) in enumerate(results):
            if record is None:
                values = [query_value, "", "", "", query_value, "", "未找到"]
            else:
                found += 1
                source = record.sheet_name
                phone = self.record_value(record, "手机号码", "号码") or query_value
                values = [
                    query_value,
                    self.record_value(record, "专页ID"),
                    self.record_value(record, "姓名", "名字"),
                    self.record_value(record, "评论贴文"),
                    phone,
                    self.record_value(record, "日期"),
                    self.corrected_source(source),
                ]
            self.query_copy_rows.append((str(values[4]), str(values[6])))
            for column, value in enumerate(values):
                self.query_table.setItem(row, column, QTableWidgetItem(str(value)))
        header = self.query_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        self.copy_query_button.setEnabled(bool(self.query_copy_rows))
        self.statusBar().showMessage(f"查询完成：输出 {len(results)} 行，匹配 {found} 行", 8000)
        self.refresh_logs()

    def copy_query_results(self) -> None:
        rows = getattr(self, "query_copy_rows", [])
        if not rows:
            QMessageBox.information(self, "没有结果", "请先查询号码。")
            return
        QApplication.clipboard().setText("\n".join(f"{phone}\t{corrected}" for phone, corrected in rows))
        self.statusBar().showMessage(f"已复制 {len(rows)} 行：号码和修正格式", 5000)

    def choose_output(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "保存提取结果", self.output_path.text(), "Excel 工作簿 (*.xlsx)")
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            self.output_path.setText(path)

    def update_output_destination(self) -> None:
        google_mode = self.output_type.currentIndex() == 1
        self.google_output_url.setEnabled(google_mode)
        for widget in getattr(self, "local_output_widgets", []):
            widget.setEnabled(not google_mode)

    def run_extract(self) -> None:
        output = self.output_path.text().strip()
        google_mode = self.output_type.currentIndex() == 1
        google_url = self.google_output_url.text().strip()
        if not google_mode and not output:
            QMessageBox.warning(self, "缺少路径", "请选择输出文件。")
            return
        if google_mode and not google_url:
            QMessageBox.warning(self, "缺少链接", "请填写目标 Google 表格链接。")
            return
        engine = DataEngine(self.store)
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        fields = split_names(self.dedup_fields.text())
        date_field = self.extract_date_field.currentText()
        direct = self.extract_mode.currentIndex() == 1
        sheet_name = self.output_sheet_name.text().strip() or "提取结果"
        self.store.set("google_output_url", google_url)
        self.store.set("google_output_sheet", sheet_name)
        self.extract_status.setText("正在提取…")
        self.run_task(
            lambda: engine.extract(
                date_field, start, end, output, fields, direct, sheet_name,
                "google" if google_mode else "local", google_url,
            ),
            self.extract_finished,
            "正在按时间提取…",
        )

    def extract_finished(self, result: dict[str, int]) -> None:
        message = f"已写入 {result['written']} 行；排除重复 {result['duplicates']} 行；无效日期 {result['invalid_dates']} 行。"
        self.extract_status.setText(message)
        self.refresh_logs()
        QMessageBox.information(self, "提取完成", message)

    def load_fields_table(self) -> None:
        schema = self.store.get("column_schema", [])
        self.resize_schema_table(self.schema_count.value())
        for row, name in enumerate(schema[:self.schema_table.rowCount()]):
            self.schema_table.setItem(row, 1, QTableWidgetItem(str(name)))
        self.update_schema_enabled(self.schema_enabled.isChecked())
        aliases = self.store.get("field_aliases", {})
        self.fields_table.setRowCount(len(aliases))
        for row, (field, names) in enumerate(aliases.items()):
            self.fields_table.setItem(row, 0, QTableWidgetItem(field))
            self.fields_table.setItem(row, 1, QTableWidgetItem(",".join(names)))

    def add_field_row(self) -> None:
        self.fields_table.insertRow(self.fields_table.rowCount())

    @staticmethod
    def excel_column(index: int) -> str:
        result = ""
        number = index + 1
        while number:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def resize_schema_table(self, count: int) -> None:
        existing = []
        if hasattr(self, "schema_table"):
            for row in range(self.schema_table.rowCount()):
                item = self.schema_table.item(row, 1)
                existing.append(item.text() if item else "")
        self.schema_table.setRowCount(count)
        for row in range(count):
            column_item = QTableWidgetItem(self.excel_column(row))
            column_item.setFlags(column_item.flags() & ~Qt.ItemIsEditable)
            self.schema_table.setItem(row, 0, column_item)
            if row < len(existing) and self.schema_table.item(row, 1) is None:
                self.schema_table.setItem(row, 1, QTableWidgetItem(existing[row]))

    def update_schema_enabled(self, enabled: bool) -> None:
        self.schema_count.setEnabled(enabled)
        self.schema_table.setEnabled(enabled)

    def fill_example_schema(self) -> None:
        names = ["专页ID", "姓名", "标签", "订阅时间", "性别", "评论贴文", "手机号码", "日期"]
        self.schema_enabled.setChecked(True)
        self.schema_count.setValue(len(names))
        for row, name in enumerate(names):
            self.schema_table.setItem(row, 1, QTableWidgetItem(name))

    def remove_field_row(self) -> None:
        if self.fields_table.currentRow() >= 0:
            self.fields_table.removeRow(self.fields_table.currentRow())

    def save_fields(self) -> None:
        aliases: dict[str, list[str]] = {}
        for row in range(self.fields_table.rowCount()):
            field_item = self.fields_table.item(row, 0)
            names_item = self.fields_table.item(row, 1)
            field = field_item.text().strip() if field_item else ""
            if field:
                aliases[field] = split_names(names_item.text() if names_item else "")
        if not aliases:
            QMessageBox.warning(self, "不能保存", "至少保留一个标准字段。")
            return
        schema = []
        for row in range(self.schema_table.rowCount()):
            item = self.schema_table.item(row, 1)
            schema.append(item.text().strip() if item else "")
        self.store.set("field_aliases", aliases)
        self.store.set("column_schema_enabled", self.schema_enabled.isChecked())
        self.store.set("column_schema", schema)
        self.refresh_field_controls()
        QMessageBox.information(self, "已保存", f"列结构和字段配置已经保存，共 {len(schema)} 列。")

    def refresh_field_controls(self) -> None:
        aliases = self.store.get("field_aliases", {})
        fields = list(aliases.keys())
        alias_lookup = {
            str(name).strip().casefold(): canonical
            for canonical, names in aliases.items()
            for name in [canonical, *names]
        }
        if self.store.get("column_schema_enabled", False):
            for header in self.store.get("column_schema", []):
                field = alias_lookup.get(str(header).strip().casefold(), str(header).strip())
                if field and field not in fields:
                    fields.append(field)
        current_query = self.query_field.currentText() if hasattr(self, "query_field") else ""
        current_date = self.extract_date_field.currentText() if hasattr(self, "extract_date_field") else ""
        for combo, current in ((self.query_field, current_query), (self.extract_date_field, current_date)):
            combo.clear()
            combo.addItems(fields)
            if current in fields:
                combo.setCurrentText(current)
        if "号码" in fields:
            self.query_field.setCurrentText("号码")
        if "日期" in fields:
            self.extract_date_field.setCurrentText("日期")

    def refresh_logs(self) -> None:
        rows = self.store.read_logs()
        self.log_table.setRowCount(len(rows))
        colors = {"ERROR": QColor("#dc2626"), "WARNING": QColor("#d97706"), "INFO": QColor("#2563eb")}
        for row_index, row in enumerate(rows):
            for column, key in enumerate(("created_at", "level", "operation", "message", "detail")):
                item = QTableWidgetItem(str(row[key]))
                if column == 1:
                    item.setForeground(colors.get(str(row[key]), QColor("#334155")))
                self.log_table.setItem(row_index, column, item)
        self.log_table.resizeColumnsToContents()
        self.log_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)

    def clear_logs(self) -> None:
        if QMessageBox.question(self, "清空日志", "确定清空全部运行日志吗？") == QMessageBox.Yes:
            self.store.clear_logs()
            self.refresh_logs()

    def choose_default_credential(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择服务账号 JSON", "", "JSON 文件 (*.json)")
        if path:
            self.default_credential.setText(path)

    def save_settings(self) -> None:
        self.store.set("global_excludes", split_names(self.global_excludes.toPlainText()))
        self.store.set("credential_path", self.default_credential.text().strip())
        QMessageBox.information(self, "已保存", "设置已经保存。")

    def run_task(self, job: Callable[[], object], success: Callable[[object], None], status: str, button: QPushButton | None = None) -> None:
        self.statusBar().showMessage(status)
        task = TaskThread(job, self)
        self.tasks.append(task)

        def done(result: object) -> None:
            self.statusBar().showMessage("操作完成", 5000)
            if button:
                button.setEnabled(True)
            success(result)
            self.tasks.remove(task)
            task.deleteLater()

        def failed(message: str) -> None:
            self.statusBar().showMessage("操作失败", 8000)
            if button:
                button.setEnabled(True)
            self.refresh_logs()
            QMessageBox.critical(self, "操作失败", message)
            self.tasks.remove(task)
            task.deleteLater()

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()


STYLE = """
QMainWindow, QWidget { background: #f7f9fc; color: #172033; font-family: "Microsoft YaHei UI"; font-size: 13px; }
#sidebar { background: #0b2748; }
#sidebar QLabel { background: transparent; }
#brand { color: white; font-size: 26px; font-weight: 700; }
#sidebarMuted { color: #9eb7d2; }
#navigation { background: transparent; border: 0; color: #cfe1f5; outline: none; }
#navigation::item { padding: 12px 13px; border-radius: 7px; margin-bottom: 4px; }
#navigation::item:selected { background: #0ea5e9; color: white; }
#navigation::item:hover { background: #16446f; }
#pageTitle { font-size: 24px; font-weight: 700; color: #10213a; }
#muted { color: #64748b; }
#card { background: white; border: 1px solid #dce5ef; border-radius: 10px; padding: 16px; }
#infoBox { background: #e8f5ff; color: #075985; border: 1px solid #b9e4ff; border-radius: 7px; padding: 11px; }
QPushButton { background: white; border: 1px solid #cbd5e1; border-radius: 7px; padding: 8px 14px; }
QPushButton:hover { border-color: #0ea5e9; color: #0369a1; }
QPushButton:disabled { color: #94a3b8; background: #e2e8f0; }
QPushButton#primary { background: #087fbb; border-color: #087fbb; color: white; font-weight: 600; }
QPushButton#primary:hover { background: #0369a1; }
QLineEdit, QTextEdit, QComboBox, QSpinBox, QDateEdit { background: white; border: 1px solid #cbd5e1; border-radius: 6px; padding: 7px; selection-background-color: #0ea5e9; }
QTableWidget { background: white; alternate-background-color: #f8fafc; border: 1px solid #dce5ef; border-radius: 8px; gridline-color: #e8edf3; }
QHeaderView::section { background: #edf3f8; color: #334155; border: 0; border-bottom: 1px solid #d7e0e9; padding: 9px; font-weight: 600; }
QCheckBox { spacing: 7px; }
QStatusBar { background: white; border-top: 1px solid #e2e8f0; color: #475569; }
"""


def create_app(icon_path: Path | None = None, data_dir: Path | None = None) -> tuple[QApplication, MainWindow]:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setOrganizationName("SheetDataHub")
    app.setStyle("Fusion")
    # 某些精简 Windows/离屏环境不会自动枚举中文字体，显式注册系统字体。
    for font_path in (Path("C:/Windows/Fonts/msyh.ttc"), Path("C:/Windows/Fonts/simsun.ttc")):
        if font_path.exists() and QFontDatabase.addApplicationFont(str(font_path)) >= 0:
            break
    app.setStyleSheet(STYLE)
    if icon_path and icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    window = MainWindow(ConfigStore(data_dir), icon_path)
    return app, window
