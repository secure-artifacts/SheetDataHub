from __future__ import annotations

import json
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Callable

from PySide6.QtCore import QDate, QThread, QTimer, Qt, Signal
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
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from .config_store import ConfigStore
from .engine import DataEngine
from .models import Record, SourceConfig
from .source_reader import (
    SourceReader,
    excel_column,
    normalize_column_schema,
    schema_field_names,
    split_names,
)
from .version import APP_VERSION, download_release_installer, fetch_latest_release, is_newer


APP_TITLE = "表数通"
QUERY_SOURCES = ["extract", "aggregate", "direct"]
COLUMN_LETTERS = [excel_column(index) for index in range(52)]
DEFAULT_QUERY_RESULT_FIELDS = ["输入电话号码", "专页ID", "姓名", "评论贴文", "号码", "日期", "修正格式"]


def query_result_headers(
    store: ConfigStore,
    source: str,
    query_field: str,
    result_fields: list[str],
) -> list[str]:
    engine = DataEngine(store)
    if source == "extract" or engine._field_kind(query_field).casefold() == "号码":
        return DEFAULT_QUERY_RESULT_FIELDS.copy()
    return result_fields or [query_field]


class ColumnMapRow(QFrame):
    changed = Signal()
    move_requested = Signal(object, int)
    remove_requested = Signal(object)

    def __init__(self, name: str = "", column: str = "A", enabled: bool = True, parent=None):
        super().__init__(parent)
        self.setObjectName("columnMapRow")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)
        self.enabled_box = QCheckBox()
        self.enabled_box.setChecked(enabled)
        self.name_edit = QLineEdit(name)
        self.name_edit.setPlaceholderText("表头名称，例如：贴文ID")
        self.column_box = QComboBox()
        self.column_box.setEditable(True)
        self.column_box.setInsertPolicy(QComboBox.NoInsert)
        self.column_box.addItems(COLUMN_LETTERS)
        self.column_box.setCurrentText((column or "A").upper())
        self.column_box.setFixedWidth(72)
        self.column_box.setObjectName("columnLetter")
        up = QToolButton()
        up.setText("↑")
        down = QToolButton()
        down.setText("↓")
        delete = QToolButton()
        delete.setText("×")
        delete.setToolTip("删除这一行")
        layout.addWidget(self.enabled_box)
        layout.addWidget(self.name_edit, 1)
        layout.addWidget(self.column_box)
        layout.addWidget(up)
        layout.addWidget(down)
        layout.addWidget(delete)
        self.enabled_box.toggled.connect(self.changed)
        self.name_edit.textChanged.connect(self.changed)
        self.column_box.currentTextChanged.connect(self.changed)
        up.clicked.connect(lambda: self.move_requested.emit(self, -1))
        down.clicked.connect(lambda: self.move_requested.emit(self, 1))
        delete.clicked.connect(lambda: self.remove_requested.emit(self))

    def value(self) -> dict[str, object]:
        return {
            "name": self.name_edit.text().strip(),
            "column": self.column_box.currentText().strip().upper() or "A",
            "enabled": self.enabled_box.isChecked(),
        }


class ColumnMapWidget(QWidget):
    changed = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)
        header = QHBoxLayout()
        enabled_label = QLabel("选用")
        enabled_label.setFixedWidth(36)
        name_label = QLabel("表头名称")
        column_label = QLabel("对应列")
        column_label.setFixedWidth(72)
        header.addWidget(enabled_label)
        header.addWidget(name_label, 1)
        header.addWidget(column_label)
        header.addSpacing(92)
        outer.addLayout(header)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setMinimumHeight(180)
        self.scroll.setMaximumHeight(280)
        self.list_host = QWidget()
        self.rows_layout = QVBoxLayout(self.list_host)
        self.rows_layout.setContentsMargins(0, 0, 0, 0)
        self.rows_layout.setSpacing(6)
        self.rows_layout.addStretch()
        self.scroll.setWidget(self.list_host)
        outer.addWidget(self.scroll)
        actions = QHBoxLayout()
        add = QPushButton("添加字段")
        add.clicked.connect(lambda: self.add_row())
        actions.addWidget(add)
        actions.addStretch()
        outer.addLayout(actions)
        self._rows: list[ColumnMapRow] = []

    def add_row(self, name: str = "", column: str = "", enabled: bool = True) -> None:
        letter = column or excel_column(len(self._rows))
        row = ColumnMapRow(name, letter, enabled, self.list_host)
        row.changed.connect(self.changed)
        row.move_requested.connect(self.move_row)
        row.remove_requested.connect(self.remove_row)
        self.rows_layout.insertWidget(self.rows_layout.count() - 1, row)
        self._rows.append(row)
        self.changed.emit()

    def move_row(self, row: ColumnMapRow, step: int) -> None:
        index = self._rows.index(row)
        target = index + step
        if target < 0 or target >= len(self._rows):
            return
        self._rows[index], self._rows[target] = self._rows[target], self._rows[index]
        self.rows_layout.removeWidget(row)
        self.rows_layout.insertWidget(target, row)
        self.changed.emit()

    def remove_row(self, row: ColumnMapRow) -> None:
        if row in self._rows:
            self._rows.remove(row)
        self.rows_layout.removeWidget(row)
        row.deleteLater()
        self.changed.emit()

    def set_schema(self, schema: list[object] | None) -> None:
        for row in list(self._rows):
            self.remove_row(row)
        entries = normalize_column_schema(schema)
        if not entries:
            self.add_row("", "A", True)
            return
        for entry in entries:
            self.add_row(str(entry.get("name") or ""), str(entry.get("column") or "A"), bool(entry.get("enabled", True)))

    def schema(self) -> list[dict[str, object]]:
        return [row.value() for row in self._rows]

    def setEnabled(self, enabled: bool) -> None:
        super().setEnabled(enabled)
        self.scroll.setEnabled(enabled)


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
        self.setMinimumWidth(760)
        self.setMinimumHeight(640)
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
        self.use_own_schema = QCheckBox("使用独立列结构（此表可指定自己的表头和对应列）")
        self.use_own_schema.setChecked(bool(source.column_schema_enabled) if source else False)
        self.schema_editor = ColumnMapWidget()
        self.schema_editor.set_schema(source.column_schema if source else [])
        form.addRow("", self.use_own_schema)
        form.addRow("字段分配", self.schema_editor)
        hint = QLabel("勾选要用的字段，填写表头名称，再选择对应的表格列（A、B、C…）。可用上下箭头调整顺序。独立列结构只作用于当前数据源。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        form.addRow("", hint)
        self.use_own_schema.toggled.connect(self.schema_editor.setEnabled)
        self.schema_editor.setEnabled(self.use_own_schema.isChecked())
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
        if self.use_own_schema.isChecked() and not schema_field_names(self.schema_editor.schema()):
            QMessageBox.warning(self, "缺少列结构", "启用独立列结构时，请至少勾选并填写一个表头名称。")
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
            column_schema_enabled=self.use_own_schema.isChecked(),
            column_schema=self.schema_editor.schema(),
        )


class MainWindow(QMainWindow):
    def __init__(self, store: ConfigStore, icon_path: Path | None = None):
        super().__init__()
        self.store = store
        self.tasks: list[TaskThread] = []
        self.icon_path = icon_path
        self._restoring_settings = True
        self.setWindowTitle(f"{APP_TITLE} · 多表数据查询与提取")
        self.resize(1220, 780)
        self.setMinimumSize(980, 650)
        if icon_path and icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self._build_ui()
        self.refresh_sources()
        self.refresh_field_controls()
        self.refresh_logs()
        self._restoring_settings = False
        QTimer.singleShot(1500, self.check_updates_silent)

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
        self.version_label = QLabel(f"v{APP_VERSION}")
        self.version_label.setObjectName("sidebarMuted")
        self.sidebar_update_button = QPushButton("检查并安装更新")
        self.sidebar_update_button.clicked.connect(self.check_updates)
        sidebar_layout.addWidget(self.version_label)
        sidebar_layout.addWidget(self.sidebar_update_button)
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
        page, layout = self._page("数据源", "配置一个或多个 Google 表格链接；每个数据源都可以有自己的列数和表头。")
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
        self.source_table = QTableWidget(0, 8)
        self.source_table.setHorizontalHeaderLabels(["启用", "名称", "表格链接", "指定 Sheet", "排除 Sheet", "表头行", "列结构", "认证"])
        self.source_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.source_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.source_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.source_table.verticalHeader().setVisible(False)
        self.source_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.source_table.doubleClicked.connect(self.edit_source)
        rule = QLabel("读取规则：指定名称留空时遍历全部；填写后仅抓取指定项；排除项始终优先。不同表格列数不一样时，在编辑数据源里勾选「独立列结构」。")
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
        page, layout = self._page(
            "数据查询",
            "切换表格来源时，查询字段会换成该表自己的表头；也可以手输列名自定义查询。",
        )
        bar = QHBoxLayout()
        self.query_field = QComboBox()
        self.query_field.setEditable(True)
        self.query_field.setInsertPolicy(QComboBox.NoInsert)
        self.query_field.setMinimumWidth(140)
        self.query_value = QTextEdit()
        self.query_value.setMaximumHeight(78)
        self.query_value.setPlaceholderText("每行一个查询值，也支持逗号分隔批量粘贴")
        self.query_mode = QComboBox()
        self.query_mode.addItems(["查询提取表", "查询汇总数据库", "直接查询数据源"])
        saved_source = str(self.store.get("query_source", "extract"))
        if saved_source in QUERY_SOURCES:
            self.query_mode.setCurrentIndex(QUERY_SOURCES.index(saved_source))
        self.query_table_label = QLabel("数据表")
        self.query_source_pick = QComboBox()
        self.query_source_pick.setMinimumWidth(140)
        self.query_fuzzy = QCheckBox("模糊匹配")
        self.query_fuzzy.setChecked(bool(self.store.get("query_fuzzy", False)))
        self.query_fuzzy.setToolTip("勾选后可用姓名、编号或部分文字查询；查询来源时也会匹配修正格式")
        self.query_date_enabled = QCheckBox("限制日期范围")
        self.query_date_enabled.setChecked(bool(self.store.get("query_date_enabled", False)))
        self.query_date_field = QComboBox()
        self.query_date_field.setEditable(True)
        self.query_date_field.setMinimumWidth(120)
        self.query_start_date = QDateEdit(QDate.currentDate().addMonths(-1))
        self.query_end_date = QDateEdit(QDate.currentDate())
        for widget, key in (
            (self.query_start_date, "query_start_date"),
            (self.query_end_date, "query_end_date"),
        ):
            widget.setCalendarPopup(True)
            widget.setDisplayFormat("yyyy-MM-dd")
            saved = QDate.fromString(str(self.store.get(key, "") or ""), "yyyy-MM-dd")
            if saved.isValid():
                widget.setDate(saved)
        self.query_mode.currentIndexChanged.connect(self.on_query_mode_changed)
        self.query_source_pick.currentIndexChanged.connect(self.on_query_table_changed)
        self.query_fuzzy.toggled.connect(self.persist_workspace_settings)
        self.query_date_enabled.toggled.connect(self.update_query_date_controls)
        self.query_date_enabled.toggled.connect(self.persist_workspace_settings)
        self.query_date_field.currentTextChanged.connect(self.persist_workspace_settings)
        self.query_start_date.dateChanged.connect(self.persist_workspace_settings)
        self.query_end_date.dateChanged.connect(self.persist_workspace_settings)
        self.query_field.currentTextChanged.connect(self.persist_workspace_settings)
        button = QPushButton("查询")
        button.setObjectName("primary")
        button.clicked.connect(self.run_query)
        self.copy_query_button = QPushButton("一键复制号码和修正格式")
        self.copy_query_button.setEnabled(False)
        self.copy_query_button.clicked.connect(self.copy_query_results)
        bar.addWidget(QLabel("表格来源"))
        bar.addWidget(self.query_mode)
        bar.addWidget(self.query_table_label)
        bar.addWidget(self.query_source_pick)
        bar.addWidget(QLabel("查询字段"))
        bar.addWidget(self.query_field)
        bar.addWidget(self.query_value, 1)
        bar.addWidget(self.query_fuzzy)
        bar.addWidget(button)
        bar.addWidget(self.copy_query_button)
        date_bar = QHBoxLayout()
        date_bar.addWidget(self.query_date_enabled)
        date_bar.addWidget(QLabel("日期字段"))
        date_bar.addWidget(self.query_date_field)
        date_bar.addWidget(QLabel("日期范围"))
        date_bar.addWidget(self.query_start_date)
        date_bar.addWidget(QLabel("至"))
        date_bar.addWidget(self.query_end_date)
        date_bar.addStretch()
        self.query_table = QTableWidget()
        self.query_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.query_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.query_table.verticalHeader().setVisible(False)
        self.query_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        hint = QLabel("提取表、汇总库、每个数据源都可以用各自表头查询。查数据源时先选数据表，查询字段会跟着切换；也可以直接输入列名。")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        self.query_result_summary = QLabel("查询结果：0 条")
        self.query_result_summary.setObjectName("muted")
        self.query_result_summary.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addLayout(bar)
        layout.addLayout(date_bar)
        layout.addWidget(hint)
        layout.addWidget(self.query_result_summary)
        layout.addWidget(self.query_table, 1)
        self.update_query_date_controls()
        return page

    def _extract_page(self) -> QWidget:
        page, layout = self._page("时间提取", "按日期范围提取数据并写入独立 Excel；历史提取记录会自动排重。")
        card = QFrame()
        card.setObjectName("card")
        form = QFormLayout(card)
        self.extract_date_field = QComboBox()
        self.extract_date_field.currentTextChanged.connect(self.persist_workspace_settings)
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
        self.dedup_fields = QLineEdit(str(self.store.get("extract_dedup_fields", "号码,日期") or "号码,日期"))
        self.dedup_fields.setPlaceholderText("多个字段用逗号分隔；留空时使用源位置和行指纹")
        self.extract_signature_enabled = QCheckBox("添加签字栏")
        self.extract_signature_enabled.setChecked(bool(self.store.get("extract_signature_enabled", False)))
        self.extract_signature_header = QLineEdit(str(self.store.get("extract_signature_header", "签字") or "签字"))
        self.extract_signature_header.setPlaceholderText("例如：签字")
        self.extract_signature_value = QLineEdit(str(self.store.get("extract_signature_value", "") or ""))
        self.extract_signature_value.setPlaceholderText("提取人姓名")
        self.extract_signature_column = QComboBox()
        self.extract_signature_column.setEditable(True)
        self.extract_signature_column.setInsertPolicy(QComboBox.NoInsert)
        self.extract_signature_column.addItems(["", *COLUMN_LETTERS])
        self.extract_signature_column.setCurrentText(str(self.store.get("extract_signature_column", "") or ""))
        signature_row = QHBoxLayout()
        signature_row.addWidget(self.extract_signature_enabled)
        signature_row.addWidget(QLabel("表头"))
        signature_row.addWidget(self.extract_signature_header)
        signature_row.addWidget(QLabel("内容"))
        signature_row.addWidget(self.extract_signature_value)
        signature_row.addWidget(QLabel("列"))
        signature_row.addWidget(self.extract_signature_column)
        self.extract_mode = QComboBox()
        self.extract_mode.addItems(["从汇总数据库提取", "直接从数据源提取"])
        self.extract_mode.setCurrentIndex(1 if str(self.store.get("extract_mode", "aggregate")) == "direct" else 0)
        self.extract_source_label = QLabel("指定数据源")
        self.extract_source_pick = QComboBox()
        self.extract_source_pick.setMinimumWidth(220)
        self.extract_source_pick.currentIndexChanged.connect(self.on_extract_source_changed)
        self.output_type = QComboBox()
        self.output_type.addItems(["本地 Excel 文件", "Google 表格链接"])
        saved_type = str(self.store.get("extract_destination_type", "local"))
        self.output_type.setCurrentIndex(1 if saved_type == "google" else 0)
        self.output_sheet_name = QLineEdit(str(self.store.get("google_output_sheet", "提取结果") or "提取结果"))
        self.output_sheet_name.setPlaceholderText("例如：查询结果")
        self.google_output_url = QLineEdit(str(self.store.get("google_output_url", "") or ""))
        self.google_output_url.setPlaceholderText("https://docs.google.com/spreadsheets/d/...")
        saved_path = str(self.store.get("extract_output_path", "") or "").strip()
        self.output_path = QLineEdit(saved_path or str(Path.home() / "Desktop" / "提取结果.xlsx"))
        choose = QPushButton("选择…")
        choose.clicked.connect(self.choose_output)
        output_row = QHBoxLayout()
        output_row.addWidget(self.output_path, 1)
        output_row.addWidget(choose)
        form.addRow("日期字段", self.extract_date_field)
        form.addRow("日期范围", date_row)
        form.addRow("去重字段", self.dedup_fields)
        form.addRow("签字栏", signature_row)
        form.addRow("数据来源", self.extract_mode)
        form.addRow(self.extract_source_label, self.extract_source_pick)
        form.addRow("输出目标", self.output_type)
        form.addRow("目标表格链接", self.google_output_url)
        form.addRow("目标工作表", self.output_sheet_name)
        form.addRow("输出文件", output_row)
        self.extract_schema_enabled = QCheckBox("提取表使用独立字段分配（表头和列可以与数据源不同）")
        self.extract_schema_enabled.setChecked(bool(self.store.get("extract_column_schema_enabled", False)))
        self.extract_schema_editor = ColumnMapWidget()
        self.extract_schema_editor.set_schema(self.store.get("extract_column_schema", []) or [])
        self.extract_schema_editor.setEnabled(self.extract_schema_enabled.isChecked())
        self.extract_schema_enabled.toggled.connect(self.extract_schema_editor.setEnabled)
        form.addRow("", self.extract_schema_enabled)
        form.addRow("提取表字段", self.extract_schema_editor)
        self.output_type.currentIndexChanged.connect(self.update_output_destination)
        self.extract_mode.currentIndexChanged.connect(self.on_extract_mode_changed)
        self.dedup_fields.editingFinished.connect(self.persist_workspace_settings)
        self.extract_signature_enabled.toggled.connect(self.persist_workspace_settings)
        self.extract_signature_header.editingFinished.connect(self.persist_workspace_settings)
        self.extract_signature_value.editingFinished.connect(self.persist_workspace_settings)
        self.extract_signature_column.currentTextChanged.connect(self.persist_workspace_settings)
        self.google_output_url.editingFinished.connect(self.persist_workspace_settings)
        self.output_sheet_name.editingFinished.connect(self.persist_workspace_settings)
        self.output_path.editingFinished.connect(self.persist_workspace_settings)
        self.extract_schema_enabled.toggled.connect(self.persist_workspace_settings)
        self.local_output_widgets = [self.output_path, choose]
        self.refresh_extract_source_picker()
        self.update_extract_source_visibility()
        self.update_output_destination()
        self.extract_button = QPushButton("开始提取")
        self.extract_button.setObjectName("primary")
        self.extract_button.clicked.connect(self.run_extract)
        self.clear_extract_cache_button = QPushButton("清除本地排重缓存")
        self.clear_extract_cache_button.clicked.connect(self.clear_extract_cache)
        extract_actions = QHBoxLayout()
        extract_actions.addWidget(self.extract_button)
        extract_actions.addWidget(self.clear_extract_cache_button)
        extract_actions.addStretch()
        self.extract_status = QLabel("等待运行")
        self.extract_status.setObjectName("muted")
        layout.addWidget(card)
        layout.addLayout(extract_actions)
        layout.addWidget(self.extract_status)
        layout.addStretch()
        return page

    def _fields_page(self) -> QWidget:
        page, layout = self._page("列与字段配置", "给每个字段勾选、起名、指定对应列。某个数据源列不一样时，到该数据源里单独分配。")
        schema_card = QFrame()
        schema_card.setObjectName("card")
        schema_layout = QVBoxLayout(schema_card)
        schema_actions = QHBoxLayout()
        self.schema_enabled = QCheckBox("启用全局字段分配")
        self.schema_enabled.setChecked(bool(self.store.get("column_schema_enabled", False)))
        example = QPushButton("填入示例")
        example.clicked.connect(self.fill_example_schema)
        schema_actions.addWidget(self.schema_enabled)
        schema_actions.addWidget(example)
        schema_actions.addStretch()
        self.schema_editor = ColumnMapWidget()
        schema_layout.addLayout(schema_actions)
        schema_layout.addWidget(self.schema_editor)
        schema_hint = QLabel("勾选字段、填写表头名称、选择对应列（A/B/C…），可用上下箭头调整顺序。已单独配置的数据源不会使用这里的分配。")
        schema_hint.setObjectName("muted")
        schema_hint.setWordWrap(True)
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
        page, layout = self._page("设置", "配置所有数据源共用的默认排除项、认证文件，以及软件更新。")
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
        version_row = QHBoxLayout()
        version_row.addWidget(QLabel(f"当前版本 v{APP_VERSION}"))
        self.update_button = QPushButton("检查并安装更新")
        self.update_button.clicked.connect(self.check_updates)
        version_row.addWidget(self.update_button)
        version_row.addStretch()
        form.addRow("软件版本", version_row)
        config_row = QHBoxLayout()
        export_config = QPushButton("导出配置")
        import_config = QPushButton("导入配置")
        export_config.clicked.connect(self.export_config_file)
        import_config.clicked.connect(self.import_config_file)
        config_row.addWidget(export_config)
        config_row.addWidget(import_config)
        config_row.addStretch()
        form.addRow("配置迁移", config_row)
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
            if source.column_schema_enabled and source.column_schema:
                schema_label = f"独立{len(schema_field_names(source.column_schema))}列"
            else:
                schema_label = "自动/全局"
            values = [
                "是" if source.enabled else "否", source.name, source.url,
                "、".join(source.include_sheets) or "全部", "、".join(source.exclude_sheets) or "—",
                str(source.header_row), schema_label, "服务账号" if source.credential_path else "公开读取",
            ]
            for column, value in enumerate(values):
                self.source_table.setItem(row, column, QTableWidgetItem(value))
        self.refresh_query_source_picker()
        self.refresh_extract_source_picker()
        if not getattr(self, "_restoring_settings", False):
            self.refresh_query_fields()

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

    def persist_workspace_settings(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        if hasattr(self, "query_mode"):
            self.store.set("query_source", QUERY_SOURCES[self.query_mode.currentIndex()])
            self.store.set("query_fuzzy", self.query_fuzzy.isChecked())
            self.store.set("query_date_enabled", self.query_date_enabled.isChecked())
            self.store.set("query_date_field", self.query_date_field.currentText().strip() or "日期")
            self.store.set("query_start_date", self.query_start_date.date().toString("yyyy-MM-dd"))
            self.store.set("query_end_date", self.query_end_date.date().toString("yyyy-MM-dd"))
            field = self.query_field.currentText().strip()
            if field:
                self.store.set("query_field", field)
                saved_map = dict(self.store.get("query_fields_by_mode", {}) or {})
                saved_map[QUERY_SOURCES[self.query_mode.currentIndex()]] = field
                self.store.set("query_fields_by_mode", saved_map)
            if hasattr(self, "query_source_pick"):
                self.store.set("query_direct_source_id", self.query_source_pick.currentData() or "")
        if hasattr(self, "output_type"):
            self.store.set("extract_destination_type", "google" if self.output_type.currentIndex() == 1 else "local")
            self.store.set("google_output_url", self.google_output_url.text().strip())
            self.store.set("google_output_sheet", self.output_sheet_name.text().strip() or "提取结果")
            self.store.set("extract_output_path", self.output_path.text().strip())
            self.store.set("extract_mode", "direct" if self.extract_mode.currentIndex() == 1 else "aggregate")
            if hasattr(self, "extract_source_pick"):
                self.store.set("extract_direct_source_id", self.extract_source_pick.currentData() or "")
            self.store.set("extract_dedup_fields", self.dedup_fields.text().strip() or "号码,日期")
            if hasattr(self, "extract_signature_enabled"):
                self.store.set("extract_signature_enabled", self.extract_signature_enabled.isChecked())
                self.store.set("extract_signature_header", self.extract_signature_header.text().strip() or "签字")
                self.store.set("extract_signature_value", self.extract_signature_value.text().strip())
                self.store.set("extract_signature_column", self.extract_signature_column.currentText().strip().upper())
            date_field = self.extract_date_field.currentText().strip()
            if date_field:
                self.store.set("extract_date_field", date_field)
            if hasattr(self, "extract_schema_enabled"):
                self.store.set("extract_column_schema_enabled", self.extract_schema_enabled.isChecked())
                self.store.set("extract_column_schema", self.extract_schema_editor.schema())

    def closeEvent(self, event) -> None:
        self.persist_workspace_settings()
        super().closeEvent(event)

    def extract_query_target(self) -> tuple[str, str]:
        sheet_name = str(self.store.get("google_output_sheet", "提取结果") or "提取结果")
        if hasattr(self, "output_sheet_name"):
            sheet_name = self.output_sheet_name.text().strip() or sheet_name
        if hasattr(self, "output_type"):
            if self.output_type.currentIndex() == 1:
                return self.google_output_url.text().strip(), sheet_name
            return self.output_path.text().strip(), sheet_name
        if str(self.store.get("extract_destination_type", "local")) == "google":
            return str(self.store.get("google_output_url", "") or ""), sheet_name
        return str(self.store.get("extract_output_path", "") or ""), sheet_name

    def current_query_mode(self) -> str:
        return QUERY_SOURCES[self.query_mode.currentIndex()]

    def current_query_source_id(self) -> str:
        if self.current_query_mode() != "direct" or not hasattr(self, "query_source_pick"):
            return ""
        return str(self.query_source_pick.currentData() or "")

    def on_query_mode_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.update_query_source_visibility()
        self.refresh_query_fields()
        self.persist_workspace_settings()

    def on_query_table_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.refresh_query_fields()
        self.persist_workspace_settings()

    def update_query_source_visibility(self) -> None:
        if not hasattr(self, "query_source_pick"):
            return
        direct = self.current_query_mode() == "direct"
        self.query_table_label.setVisible(direct)
        self.query_source_pick.setVisible(direct)

    def current_extract_source_id(self) -> str:
        if not hasattr(self, "extract_source_pick"):
            return ""
        if not hasattr(self, "extract_mode") or self.extract_mode.currentIndex() != 1:
            return ""
        return str(self.extract_source_pick.currentData() or "")

    def update_extract_source_visibility(self) -> None:
        if not hasattr(self, "extract_source_pick"):
            return
        direct = self.extract_mode.currentIndex() == 1
        self.extract_source_label.setVisible(direct)
        self.extract_source_pick.setVisible(direct)

    def refresh_extract_source_picker(self) -> None:
        if not hasattr(self, "extract_source_pick"):
            return
        restoring = self._restoring_settings
        self._restoring_settings = True
        saved = str(self.store.get("extract_direct_source_id", "") or "")
        current = self.extract_source_pick.currentData()
        self.extract_source_pick.clear()
        self.extract_source_pick.addItem("全部数据源", "")
        for source in self.store.load_sources():
            if source.enabled:
                self.extract_source_pick.addItem(source.name, source.id)
        target = current if current not in (None, "") else saved
        index = self.extract_source_pick.findData(target)
        self.extract_source_pick.setCurrentIndex(index if index >= 0 else 0)
        self._restoring_settings = restoring

    def update_query_date_controls(self) -> None:
        enabled = self.query_date_enabled.isChecked()
        for widget in (self.query_date_field, self.query_start_date, self.query_end_date):
            widget.setEnabled(enabled)

    def refresh_query_source_picker(self) -> None:
        if not hasattr(self, "query_source_pick"):
            return
        restoring = self._restoring_settings
        self._restoring_settings = True
        saved = str(self.store.get("query_direct_source_id", "") or "")
        current = self.query_source_pick.currentData()
        self.query_source_pick.clear()
        self.query_source_pick.addItem("全部数据源", "")
        for source in self.store.load_sources():
            if source.enabled:
                self.query_source_pick.addItem(source.name, source.id)
        target = current if current not in (None, "") else saved
        index = self.query_source_pick.findData(target)
        self.query_source_pick.setCurrentIndex(index if index >= 0 else 0)
        self._restoring_settings = restoring
        self.update_query_source_visibility()

    def refresh_query_fields(self) -> None:
        if not hasattr(self, "query_field"):
            return
        mode = self.current_query_mode()
        extract_target, extract_sheet = self.extract_query_target()
        fields = DataEngine(self.store).list_query_fields(
            mode, extract_target, extract_sheet, self.current_query_source_id(),
        )
        restoring = self._restoring_settings
        self._restoring_settings = True
        current = self.query_field.currentText().strip()
        saved_map = self.store.get("query_fields_by_mode", {}) or {}
        saved = str(saved_map.get(mode) or self.store.get("query_field", "") or "").strip()
        self.query_field.clear()
        self.query_field.addItems(fields)
        if saved in fields:
            pick = saved
        elif current in fields:
            pick = current
        elif "号码" in fields:
            pick = "号码"
        elif "手机号码" in fields:
            pick = "手机号码"
        elif fields:
            pick = fields[0]
        else:
            pick = saved or current
        if pick:
            if pick not in fields:
                self.query_field.insertItem(0, pick)
            self.query_field.setCurrentText(pick)
        current_date_field = self.query_date_field.currentText().strip()
        saved_date_field = str(self.store.get("query_date_field", "日期") or "日期").strip()
        self.query_date_field.clear()
        self.query_date_field.addItems(fields)
        date_pick = ""
        for candidate in (current_date_field, saved_date_field):
            if candidate in fields:
                date_pick = candidate
                break
        if not date_pick:
            date_pick = DataEngine(self.store).suggest_field(fields, "日期")
        if date_pick:
            self.query_date_field.setCurrentText(date_pick)
        self._restoring_settings = restoring

    def run_query(self) -> None:
        values = split_names(self.query_value.toPlainText())
        if not values:
            QMessageBox.warning(self, "请输入", "请输入一个或多个查询号码。")
            return
        engine = DataEngine(self.store)
        field = self.query_field.currentText().strip()
        if not field:
            QMessageBox.warning(self, "请选择", "请选择或输入要查询的字段。")
            return
        source = self.current_query_mode()
        exact = not self.query_fuzzy.isChecked()
        extract_target, extract_sheet = self.extract_query_target()
        source_id = self.current_query_source_id()
        date_field = self.query_date_field.currentText().strip() if self.query_date_enabled.isChecked() else ""
        start_date = self.query_start_date.date().toPython() if date_field else None
        end_date = self.query_end_date.date().toPython() if date_field else None
        if source == "extract" and not extract_target:
            QMessageBox.warning(self, "缺少提取表", "请先在「时间提取」页填写目标 Google 表格链接或本地输出文件。")
            return
        self.persist_workspace_settings()
        result_fields = query_result_headers(
            self.store,
            source,
            field,
            engine.list_query_fields(source, extract_target, extract_sheet, source_id),
        )
        if hasattr(self, "query_result_summary"):
            self.query_result_summary.setText(f"正在查询：输入 {len(values)} 个值…")
        self.run_task(
            lambda: engine.query_many(
                field, values, source == "direct", exact, source, extract_target, extract_sheet, source_id,
                date_field, start_date, end_date,
            ),
            lambda results: self.show_query_results(results, result_fields, field),
            "正在批量查询…",
        )

    @staticmethod
    def corrected_source(source: str) -> str:
        return DataEngine.corrected_source(source)

    @staticmethod
    def record_value(record: Record, *names: str) -> str:
        for name in names:
            value = record.values.get(name, "")
            if value:
                return value
        return ""

    def show_query_results(
        self,
        results: list[tuple[str, Record | None]],
        fields: list[str] | None = None,
        query_field: str = "号码",
    ) -> None:
        if not fields:
            engine = DataEngine(self.store)
            if engine._field_kind(query_field).casefold() == "号码":
                fields = DEFAULT_QUERY_RESULT_FIELDS.copy()
            else:
                fields = [query_field]
        self.query_table.setColumnCount(len(fields))
        self.query_table.setHorizontalHeaderLabels(fields)
        self.query_table.setRowCount(len(results))
        self.query_copy_rows: list[tuple[str, str]] = []
        engine = DataEngine(self.store)
        query_canonical = engine._canonical_field(query_field).casefold()
        found = 0
        for row, (query_value, record) in enumerate(results):
            if record is None:
                values = ["" for _ in fields]
                for column, header_name in enumerate(fields):
                    if header_name.strip() in {"输入电话号码", "输入电话", "查询值"}:
                        values[column] = query_value
                    elif engine._canonical_field(header_name).casefold() == query_canonical:
                        values[column] = query_value
                        break
                phone = query_value if query_canonical == "号码" else ""
                corrected = "未找到"
            else:
                found += 1
                values = [
                    query_value if header_name.strip() in {"输入电话号码", "输入电话", "查询值"}
                    else engine.query_display_value(record, header_name)
                    for header_name in fields
                ]
                phone = engine._field_value(record, "号码") or (query_value if query_canonical == "号码" else "")
                corrected = self.corrected_source(record.sheet_name)
            self.query_copy_rows.append((str(phone), str(corrected)))
            for column, value in enumerate(values):
                self.query_table.setItem(row, column, QTableWidgetItem(str(value)))
        header = self.query_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        for column, header_name in enumerate(fields):
            if header_name.strip().casefold() in {"评论贴文", "链接", "网址", "url", "link"}:
                header.setSectionResizeMode(column, QHeaderView.Stretch)
        self.copy_query_button.setEnabled(bool(self.query_copy_rows))
        missing = len(results) - found
        input_count = len({str(query_value) for query_value, _ in results})
        summary = f"查询结果：共 {len(results)} 条；匹配 {found} 条；未找到 {missing} 条；输入 {input_count} 个值"
        if hasattr(self, "query_result_summary"):
            self.query_result_summary.setText(summary)
        message = f"查询完成：共 {len(results)} 条，匹配 {found} 条"
        if self.query_mode.currentIndex() == 0 and missing:
            message += "；未找到的可把表格来源改成「直接查询数据源」再查一次"
        self.statusBar().showMessage(message, 8000)
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
            self.persist_workspace_settings()

    def update_output_destination(self) -> None:
        google_mode = self.output_type.currentIndex() == 1
        self.google_output_url.setEnabled(google_mode)
        for widget in getattr(self, "local_output_widgets", []):
            widget.setEnabled(not google_mode)
        self.persist_workspace_settings()

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
        self.persist_workspace_settings()
        engine = DataEngine(self.store)
        start = self.start_date.date().toPython()
        end = self.end_date.date().toPython()
        fields = split_names(self.dedup_fields.text())
        date_field = self.extract_date_field.currentText()
        direct = self.extract_mode.currentIndex() == 1
        source_id = self.current_extract_source_id()
        sheet_name = self.output_sheet_name.text().strip() or "提取结果"
        self.store.set("google_output_url", google_url)
        self.store.set("google_output_sheet", sheet_name)
        self.store.set("extract_destination_type", "google" if google_mode else "local")
        self.store.set("extract_output_path", output)
        self.extract_status.setText("正在提取…")
        self.run_task(
            lambda: engine.extract(
                date_field, start, end, output, fields, direct, sheet_name,
                "google" if google_mode else "local", google_url, source_id,
            ),
            self.extract_finished,
            "正在按时间提取…",
            self.extract_button,
            self.extract_failed,
        )

    def clear_extract_cache(self) -> None:
        total = self.store.count_extracted()
        if total <= 0:
            QMessageBox.information(self, "本地排重缓存", "当前没有本地提取排重缓存。")
            self.extract_status.setText("本地排重缓存为空。")
            return
        answer = QMessageBox.question(
            self,
            "清除本地排重缓存",
            f"确定清除 {total} 条本地提取排重缓存吗？\n\n"
            "这不会删除数据源、字段配置、汇总库或目标表格数据。清除后仍会根据目标提取表已有行继续排重。",
        )
        if answer != QMessageBox.Yes:
            return
        removed = self.store.clear_extracted()
        self.store.log("INFO", "时间提取", f"已清除本地提取排重缓存 {removed} 条")
        self.extract_status.setText(f"已清除本地排重缓存：{removed} 条。")
        self.refresh_logs()
        QMessageBox.information(self, "清除完成", f"已清除本地提取排重缓存 {removed} 条。")

    def extract_finished(self, result: dict[str, int]) -> None:
        message = f"提取已结束：已写入 {result['written']} 行；排除重复 {result['duplicates']} 行；无效日期 {result['invalid_dates']} 行。"
        self.extract_status.setText(message)
        self.refresh_logs()
        QMessageBox.information(self, "提取完成", message)

    def extract_failed(self, message: str) -> None:
        self.extract_status.setText(f"提取已结束：失败。{message}")

    def load_fields_table(self) -> None:
        self.schema_editor.set_schema(self.store.get("column_schema", []))
        self.update_schema_enabled(self.schema_enabled.isChecked())
        aliases = self.store.get("field_aliases", {})
        self.fields_table.setRowCount(len(aliases))
        for row, (field, names) in enumerate(aliases.items()):
            self.fields_table.setItem(row, 0, QTableWidgetItem(field))
            self.fields_table.setItem(row, 1, QTableWidgetItem(",".join(names)))

    def add_field_row(self) -> None:
        self.fields_table.insertRow(self.fields_table.rowCount())

    def update_schema_enabled(self, enabled: bool) -> None:
        self.schema_editor.setEnabled(enabled)

    def fill_example_schema(self) -> None:
        names = ["专页ID", "姓名", "标签", "订阅时间", "性别", "评论贴文", "手机号码", "日期"]
        self.schema_enabled.setChecked(True)
        self.schema_editor.set_schema([
            {"name": name, "column": excel_column(index), "enabled": True}
            for index, name in enumerate(names)
        ])

    def remove_field_row(self) -> None:
        if self.fields_table.currentRow() >= 0:
            self.fields_table.removeRow(self.fields_table.currentRow())

    def save_fields(self) -> None:
        if not self.save_fields_to_store():
            return
        self.refresh_field_controls()
        QMessageBox.information(self, "已保存", f"字段分配已经保存，共 {len(schema_field_names(self.schema_editor.schema()))} 个启用字段。")

    def refresh_field_controls(self) -> None:
        aliases = self.store.get("field_aliases", {})
        mode = "direct" if hasattr(self, "extract_mode") and self.extract_mode.currentIndex() == 1 else "aggregate"
        source_id = self.current_extract_source_id() if mode == "direct" else ""
        fields = DataEngine(self.store).list_query_fields(mode, source_id=source_id) or list(aliases.keys())
        restoring = self._restoring_settings
        self._restoring_settings = True
        current_date = self.extract_date_field.currentText() if hasattr(self, "extract_date_field") else ""
        saved_date = str(self.store.get("extract_date_field", "") or "")
        if hasattr(self, "extract_date_field"):
            self.extract_date_field.clear()
            self.extract_date_field.addItems(fields)
            date_field = current_date if current_date in fields else (saved_date if saved_date in fields else "")
            if not date_field:
                date_field = DataEngine(self.store).suggest_field(fields, "日期") or (fields[0] if fields else "")
            if date_field:
                self.extract_date_field.setCurrentText(date_field)
        self._restoring_settings = restoring
        self.refresh_query_source_picker()
        self.refresh_query_fields()

    def on_extract_mode_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.update_extract_source_visibility()
        self.refresh_field_controls()
        self.persist_workspace_settings()

    def on_extract_source_changed(self) -> None:
        if getattr(self, "_restoring_settings", False):
            return
        self.refresh_field_controls()
        self.persist_workspace_settings()

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

    def check_updates_silent(self) -> None:
        task = TaskThread(fetch_latest_release, self)
        self.tasks.append(task)

        def done(info: object) -> None:
            self.tasks.remove(task)
            task.deleteLater()
            if not isinstance(info, dict):
                return
            if is_newer(str(info.get("version") or ""), APP_VERSION):
                self.statusBar().showMessage(
                    f"发现新版本 v{info['version']}，可点击「检查并安装更新」直接安装",
                    20000,
                )

        def failed(_message: str) -> None:
            if task in self.tasks:
                self.tasks.remove(task)
            task.deleteLater()

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()

    def _set_update_buttons_enabled(self, enabled: bool) -> None:
        self.sidebar_update_button.setEnabled(enabled)
        if hasattr(self, "update_button"):
            self.update_button.setEnabled(enabled)

    def check_updates(self) -> None:
        self._set_update_buttons_enabled(False)
        self.statusBar().showMessage("正在检查更新…")
        task = TaskThread(fetch_latest_release, self)
        self.tasks.append(task)

        def done(info: object) -> None:
            self._set_update_buttons_enabled(True)
            self.tasks.remove(task)
            task.deleteLater()
            if isinstance(info, dict):
                self.show_update_result(info)

        def failed(message: str) -> None:
            self._set_update_buttons_enabled(True)
            self.statusBar().showMessage("检查更新失败", 8000)
            if task in self.tasks:
                self.tasks.remove(task)
            task.deleteLater()
            QMessageBox.critical(self, "检查更新失败", message)

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()

    def show_update_result(self, info: dict[str, str]) -> None:
        self._set_update_buttons_enabled(True)
        latest = str(info.get("version") or "")
        if is_newer(latest, APP_VERSION):
            box = QMessageBox(self)
            box.setWindowTitle("发现新版本")
            box.setText(f"当前版本：v{APP_VERSION}\n最新版本：v{latest}")
            box.setInformativeText("软件将自动下载安装包，随后关闭当前版本并启动安装。")
            install_button = box.addButton("立即下载安装", QMessageBox.AcceptRole)
            box.addButton("稍后", QMessageBox.RejectRole)
            box.exec()
            if box.clickedButton() is install_button:
                self.download_and_install_update(info)
            return
        QMessageBox.information(self, "已是最新", f"当前已经是最新版本 v{APP_VERSION}。")

    def download_and_install_update(self, info: dict[str, str]) -> None:
        self._set_update_buttons_enabled(False)
        self.statusBar().showMessage("正在下载安装包，请稍候…")
        task = TaskThread(lambda: download_release_installer(info, self.store.data_dir), self)
        self.tasks.append(task)

        def done(path: object) -> None:
            self._set_update_buttons_enabled(True)
            self.tasks.remove(task)
            task.deleteLater()
            installer = Path(str(path))
            QMessageBox.information(self, "下载完成", "安装包已下载，将关闭当前软件并启动安装程序。")
            subprocess.Popen([str(installer)], cwd=str(installer.parent))
            QApplication.quit()

        def failed(message: str) -> None:
            self._set_update_buttons_enabled(True)
            if task in self.tasks:
                self.tasks.remove(task)
            task.deleteLater()
            self.statusBar().showMessage("更新下载失败", 8000)
            QMessageBox.critical(self, "更新失败", message)

        task.succeeded.connect(done)
        task.failed.connect(failed)
        task.start()

    def choose_default_credential(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "选择服务账号 JSON", "", "JSON 文件 (*.json)")
        if path:
            self.default_credential.setText(path)

    def save_settings(self) -> None:
        self.store.set("global_excludes", split_names(self.global_excludes.toPlainText()))
        self.store.set("credential_path", self.default_credential.text().strip())
        QMessageBox.information(self, "已保存", "设置已经保存。")

    def export_config_file(self) -> None:
        self.persist_workspace_settings()
        self.save_fields_to_store(silent=True)
        default_path = str(Path.home() / "Desktop" / "表数通配置.json")
        path, _ = QFileDialog.getSaveFileName(self, "导出配置", default_path, "JSON 配置文件 (*.json)")
        if not path:
            return
        if not path.lower().endswith(".json"):
            path += ".json"
        payload = self.store.export_config()
        Path(path).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self.store.log("INFO", "配置迁移", f"已导出配置：{path}")
        self.refresh_logs()
        QMessageBox.information(self, "导出完成", f"配置已导出：\n{path}")

    def import_config_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "导入配置", "", "JSON 配置文件 (*.json)")
        if not path:
            return
        answer = QMessageBox.question(
            self,
            "导入配置",
            "导入后会覆盖当前数据源、字段映射、列配置、查询和提取设置。\n\n"
            "不会删除本地排重缓存、运行日志或汇总数据库。确定继续吗？",
        )
        if answer != QMessageBox.Yes:
            return
        try:
            payload = json.loads(Path(path).read_text(encoding="utf-8"))
            settings_count, source_count = self.store.import_config(payload)
        except Exception as exc:
            QMessageBox.critical(self, "导入失败", str(exc))
            return
        self.store.log("INFO", "配置迁移", f"已导入配置：{path}，设置 {settings_count} 项，数据源 {source_count} 个")
        self.apply_store_to_ui()
        self.refresh_logs()
        QMessageBox.information(
            self,
            "导入完成",
            f"已导入配置：{source_count} 个数据源，{settings_count} 项设置。\n"
            "如果服务账号 JSON 路径在这台电脑不存在，请到设置里重新选择。",
        )

    def save_fields_to_store(self, silent: bool = False) -> bool:
        if not hasattr(self, "fields_table"):
            return True
        aliases: dict[str, list[str]] = {}
        for row in range(self.fields_table.rowCount()):
            field_item = self.fields_table.item(row, 0)
            names_item = self.fields_table.item(row, 1)
            field = field_item.text().strip() if field_item else ""
            if not field:
                continue
            aliases[field] = split_names(names_item.text() if names_item else "")
        if not aliases:
            if not silent:
                QMessageBox.warning(self, "不能保存", "至少保留一个标准字段。")
            return False
        schema = self.schema_editor.schema()
        self.store.set("field_aliases", aliases)
        self.store.set("column_schema_enabled", self.schema_enabled.isChecked())
        self.store.set("column_schema", schema)
        return True

    def apply_store_to_ui(self) -> None:
        restoring = self._restoring_settings
        self._restoring_settings = True
        try:
            if hasattr(self, "global_excludes"):
                self.global_excludes.setPlainText("\n".join(self.store.get("global_excludes", [])))
                self.default_credential.setText(str(self.store.get("credential_path", "") or ""))
            if hasattr(self, "query_mode"):
                saved_source = str(self.store.get("query_source", "extract") or "extract")
                if saved_source in QUERY_SOURCES:
                    self.query_mode.setCurrentIndex(QUERY_SOURCES.index(saved_source))
                self.query_fuzzy.setChecked(bool(self.store.get("query_fuzzy", False)))
                self.query_date_enabled.setChecked(bool(self.store.get("query_date_enabled", False)))
                for widget, key in (
                    (self.query_start_date, "query_start_date"),
                    (self.query_end_date, "query_end_date"),
                ):
                    saved = QDate.fromString(str(self.store.get(key, "") or ""), "yyyy-MM-dd")
                    if saved.isValid():
                        widget.setDate(saved)
            if hasattr(self, "output_type"):
                saved_type = str(self.store.get("extract_destination_type", "local") or "local")
                self.output_type.setCurrentIndex(1 if saved_type == "google" else 0)
                self.google_output_url.setText(str(self.store.get("google_output_url", "") or ""))
                self.output_sheet_name.setText(str(self.store.get("google_output_sheet", "提取结果") or "提取结果"))
                self.output_path.setText(str(self.store.get("extract_output_path", "") or ""))
                self.extract_mode.setCurrentIndex(1 if str(self.store.get("extract_mode", "aggregate")) == "direct" else 0)
                self.dedup_fields.setText(str(self.store.get("extract_dedup_fields", "号码,日期") or "号码,日期"))
                self.extract_signature_enabled.setChecked(bool(self.store.get("extract_signature_enabled", False)))
                self.extract_signature_header.setText(str(self.store.get("extract_signature_header", "签字") or "签字"))
                self.extract_signature_value.setText(str(self.store.get("extract_signature_value", "") or ""))
                self.extract_signature_column.setCurrentText(str(self.store.get("extract_signature_column", "") or ""))
                self.extract_schema_enabled.setChecked(bool(self.store.get("extract_column_schema_enabled", False)))
                self.extract_schema_editor.set_schema(self.store.get("extract_column_schema", []) or [])
                self.extract_schema_editor.setEnabled(self.extract_schema_enabled.isChecked())
            if hasattr(self, "schema_enabled"):
                self.schema_enabled.setChecked(bool(self.store.get("column_schema_enabled", False)))
                self.schema_editor.set_schema(self.store.get("column_schema", []) or [])
                self.update_schema_enabled(self.schema_enabled.isChecked())
                self.load_fields_table()
            self.refresh_sources()
            self.refresh_extract_source_picker()
            self.refresh_query_source_picker()
            self.update_output_destination()
            self.update_extract_source_visibility()
            self.update_query_date_controls()
            self.refresh_field_controls()
        finally:
            self._restoring_settings = restoring

    def run_task(
        self,
        job: Callable[[], object],
        success: Callable[[object], None],
        status: str,
        button: QPushButton | None = None,
        failure: Callable[[str], None] | None = None,
    ) -> None:
        self.statusBar().showMessage(status)
        if button:
            button.setEnabled(False)
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
            if failure:
                failure(message)
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
#sidebar QPushButton { background: transparent; border: 1px solid #3d6a94; color: #cfe1f5; padding: 6px 10px; }
#sidebar QPushButton:hover { border-color: #0ea5e9; color: white; }
#navigation { background: transparent; border: 0; color: #cfe1f5; outline: none; }
#navigation::item { padding: 12px 13px; border-radius: 7px; margin-bottom: 4px; }
#navigation::item:selected { background: #0ea5e9; color: white; }
#navigation::item:hover { background: #16446f; }
#pageTitle { font-size: 24px; font-weight: 700; color: #10213a; }
#muted { color: #64748b; }
#card { background: white; border: 1px solid #dce5ef; border-radius: 10px; padding: 16px; }
#columnMapRow { background: white; border: 1px solid #e2eaf2; border-radius: 8px; }
#columnLetter { font-weight: 600; }
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
    app.setApplicationVersion(APP_VERSION)
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
