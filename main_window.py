from __future__ import annotations

from pathlib import Path

import pandas as pd
from qtpy.QtCore import QPoint, QSettings, Qt, Signal
from qtpy.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QSpinBox,
    QStatusBar,
    QTableView,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QToolBar,
    QVBoxLayout,
    QWidget,
)
from qtpy.QtGui import QBrush, QColor, QFont

from dataframe_model import PandasModel
from excel_service import (
    ExcelServiceError,
    apply_column_rules,
    clean_dataframe,
    export_dataframe,
    load_prepared_dataframe,
    merge_files,
    split_and_export,
)
from project_config import ColumnRule, ProjectConfig


PROJECT_FILTER = "PyPandas 配置 (*.pypandas.json)"
DATA_FILTER = "表格文件 (*.xlsx *.xls *.xlsm *.csv)"


class _CheckPopup(QFrame):
    """多选弹出列表 —— 包含复选框，点击条目切换选中。"""

    closed = Signal()
    checked_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent, Qt.WindowType.Popup)
        self.setFrameStyle(QFrame.Shadow.Plain | QFrame.Shape.Box)
        self.setLineWidth(1)

        self._list = QListWidget(self)
        self._list.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self._list.itemClicked.connect(self._on_item_clicked)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._list)
        self.setLayout(layout)

        self._data_map: dict[int, str] = {}  # item index → original_name

    def populate(self, items: list[tuple[str, str]], checked: set[str]) -> None:
        """填充 (显示名, 原始名) 列表。"""
        self._list.blockSignals(True)
        self._list.clear()
        self._data_map.clear()
        for row, (display_name, original_name) in enumerate(items):
            list_item = QListWidgetItem(display_name)
            list_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsSelectable
            )
            list_item.setCheckState(
                Qt.CheckState.Checked
                if original_name in checked
                else Qt.CheckState.Unchecked
            )
            self._list.addItem(list_item)
            self._data_map[row] = original_name
        self._list.blockSignals(False)

    def checked_originals(self) -> list[str]:
        result: list[str] = []
        for row in range(self._list.count()):
            item = self._list.item(row)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                original = self._data_map.get(row)
                if original:
                    result.append(original)
        return result

    def _on_item_clicked(self, item: QListWidgetItem) -> None:
        """切换复选框 —— checkState 已由 QListWidget 自动切换，通知父控件。"""
        self.checked_changed.emit()


class CheckableComboBox(QWidget):
    """由按钮 + 弹出多选列表组成的下拉选择控件。"""

    selection_changed = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._button = QPushButton("选择拆分列（可多选）", self)
        self._button.setStyleSheet("text-align: left;")
        self._button.clicked.connect(self._toggle_popup)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self._button)
        self.setLayout(layout)

        self._popup: _CheckPopup | None = None
        self._items: list[tuple[str, str]] = []  # (display, original)
        self._checked_originals: set[str] = set()

    def populate_items(self, items: list[tuple[str, str]], checked: set[str]) -> None:
        """重新填充条目并更新显示。"""
        self._items = items
        # 仅保留仍然存在的列
        valid = {orig for _, orig in items}
        self._checked_originals &= valid
        if checked:
            self._checked_originals = set(checked) & valid
        if self._popup is not None:
            self._popup.populate(items, self._checked_originals)
        self._update_button_text()

    def checked_items(self) -> list[str]:
        return list(self._checked_originals)

    def _update_button_text(self) -> None:
        if not self._checked_originals or not self._items:
            self._button.setText("选择拆分列（可多选）")
            return
        names: list[str] = []
        for display, original in self._items:
            if original in self._checked_originals:
                names.append(display)
        if names:
            self._button.setText("、".join(names))
        else:
            self._button.setText("选择拆分列（可多选）")

    def _toggle_popup(self) -> None:
        if self._popup is not None:
            self._close_popup()
            return

        self._popup = _CheckPopup(self.window())
        self._popup.populate(self._items, self._checked_originals)
        self._popup.checked_changed.connect(self._on_check_changed)
        self._popup.closed.connect(self._on_popup_closed)

        btn_rect = self._button.rect()
        global_pos = self._button.mapToGlobal(QPoint(0, btn_rect.height()))
        self._popup.move(global_pos)
        self._popup.setMinimumWidth(self._button.width())
        self._popup.show()

    def _on_check_changed(self) -> None:
        """弹窗内复选框变化 → 实时同步选中状态并更新按钮文案。"""
        if self._popup is not None:
            self._checked_originals = set(self._popup.checked_originals())
            self._update_button_text()

    def _close_popup(self) -> None:
        if self._popup is not None:
            # 收起前从弹窗拉取最新选中状态
            self._checked_originals = set(self._popup.checked_originals())
            self._popup.close()
            self._popup.deleteLater()
            self._popup = None
            self._update_button_text()
            self.selection_changed.emit()

    def _on_popup_closed(self) -> None:
        self._close_popup()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.settings = QSettings("墨韩", "墨韩表格工具箱")
        self.project_config = ProjectConfig()
        self.project_path: str = ""
        self.current_original_df = pd.DataFrame()
        self.current_view_df = pd.DataFrame()
        self._updating_columns_table = False
        self._building_recent_menus = False

        self.setWindowTitle("墨韩表格工具箱")
        self.resize(1200, 700)

        self._build_ui()
        self._create_actions()
        self._populate_recent_menus()
        self._load_settings_into_form()
        self._update_ui_state()

    def _build_ui(self) -> None:
        self._build_left_panel()
        self._build_right_panel()

        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_splitter.addWidget(self.left_panel)
        main_splitter.addWidget(self.right_panel)
        main_splitter.setStretchFactor(0, 0)
        main_splitter.setStretchFactor(1, 1)
        main_splitter.setSizes([360, 1000])
        self.setCentralWidget(main_splitter)

        self.status_bar = QStatusBar(self)
        self.setStatusBar(self.status_bar)

    def _build_left_panel(self) -> None:
        self.left_panel = QWidget()
        layout = QVBoxLayout(self.left_panel)

        source_group = QGroupBox("配置文件")
        source_layout = QVBoxLayout(source_group)
        self.file_list = QListWidget()
        self.file_list.itemSelectionChanged.connect(self._handle_file_selection_changed)
        source_layout.addWidget(self.file_list)

        source_buttons = QHBoxLayout()
        add_files_button = QPushButton("添加文件")
        add_files_button.clicked.connect(self.open_files)
        remove_file_button = QPushButton("移除选中")
        remove_file_button.clicked.connect(self.remove_selected_files)
        source_buttons.addWidget(add_files_button)
        source_buttons.addWidget(remove_file_button)
        source_layout.addLayout(source_buttons)
        layout.addWidget(source_group)

        settings_group = QGroupBox("处理参数")
        settings_layout = QFormLayout(settings_group)

        self.sheet_name_edit = QLineEdit()
        self.sheet_name_edit.setPlaceholderText("留空表示第一个工作表")
        settings_layout.addRow("工作表名", self.sheet_name_edit)

        self.header_row_spin = QSpinBox()
        self.header_row_spin.setMinimum(1)
        self.header_row_spin.setMaximum(9999)
        self.header_row_spin.setValue(1)
        self.header_row_spin.valueChanged.connect(self._handle_header_row_changed)
        settings_layout.addRow("标题所在行", self.header_row_spin)

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setPlaceholderText("拆分导出的目录")
        choose_output_dir_button = QPushButton("选择")
        choose_output_dir_button.clicked.connect(self.choose_output_dir)
        output_dir_layout = QHBoxLayout()
        output_dir_layout.addWidget(self.output_dir_edit)
        output_dir_layout.addWidget(choose_output_dir_button)
        settings_layout.addRow("输出目录", _wrap_layout(output_dir_layout))

        self.output_prefix_edit = QLineEdit("新")
        settings_layout.addRow("输出前缀", self.output_prefix_edit)

        self.export_format_combo = QComboBox()
        self.export_format_combo.addItems(["xlsx", "csv", "xlsm"])
        settings_layout.addRow("导出格式", self.export_format_combo)

        self.split_column_combo = CheckableComboBox()
        settings_layout.addRow("拆分列", self.split_column_combo)

        self.split_values_edit = QPlainTextEdit()
        self.split_values_edit.setPlaceholderText("可选：只导出这些值，每行一个")
        self.split_values_edit.setFixedHeight(90)
        settings_layout.addRow("拆分值过滤", self.split_values_edit)

        self.sort_column_edit = QLineEdit()
        self.sort_column_edit.setPlaceholderText("可选：排序列")
        settings_layout.addRow("排序列", self.sort_column_edit)

        self.remove_empty_rows_check = QCheckBox("移除全空行")
        self.remove_empty_rows_check.setChecked(True)
        settings_layout.addRow("", self.remove_empty_rows_check)

        self.drop_duplicates_check = QCheckBox("去除重复行")
        settings_layout.addRow("", self.drop_duplicates_check)

        apply_button = QPushButton("刷新预览")
        apply_button.clicked.connect(self.reload_active_file)
        settings_layout.addRow("", apply_button)

        layout.addWidget(settings_group)
        layout.addStretch(1)

    def _build_right_panel(self) -> None:
        self.right_panel = QWidget()
        layout = QVBoxLayout(self.right_panel)

        self.summary_label = QLabel("未加载数据")
        layout.addWidget(self.summary_label)

        preview_tabs = QTabWidget()
        layout.addWidget(preview_tabs, 1)

        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        self.table_view = QTableView()
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table_view.horizontalHeader().setStretchLastSection(True)
        self.table_model = PandasModel()
        self.table_view.setModel(self.table_model)
        preview_layout.addWidget(self.table_view)
        preview_tabs.addTab(preview_widget, "数据预览")

        self.columns_table = QTableWidget(0, 3)
        self.columns_table.setHorizontalHeaderLabels(["保留", "原列名", "新列名"])
        self.columns_table.setColumnWidth(0, 64)
        self.columns_table.setColumnWidth(1, 180)
        self.columns_table.setColumnWidth(2, 180)
        self.columns_table.horizontalHeader().setStretchLastSection(True)
        self.columns_table.verticalHeader().setVisible(False)
        self.columns_table.itemChanged.connect(self._handle_columns_table_changed)
        preview_tabs.addTab(self.columns_table, "列配置")

        self.log_edit = QPlainTextEdit()
        self.log_edit.setReadOnly(True)
        preview_tabs.addTab(self.log_edit, "日志")

        buttons_layout = QHBoxLayout()
        save_current_button = QPushButton("导出当前结果")
        save_current_button.clicked.connect(self.save_current_file_as)
        split_export_button = QPushButton("按列拆分导出")
        split_export_button.clicked.connect(self.export_split_files)
        merge_button = QPushButton("合并多个表格")
        merge_button.clicked.connect(self.merge_project_files)
        buttons_layout.addWidget(save_current_button)
        buttons_layout.addWidget(split_export_button)
        buttons_layout.addWidget(merge_button)
        buttons_layout.addStretch(1)
        layout.addLayout(buttons_layout)

    def _create_actions(self) -> None:
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("文件")
        project_menu = menu_bar.addMenu("配置")
        tools_menu = menu_bar.addMenu("工具")
        help_menu = menu_bar.addMenu("帮助")

        toolbar = QToolBar("主工具栏", self)
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        self.new_project_action = file_menu.addAction("新建配置")
        self.new_project_action.triggered.connect(self.new_project)
        toolbar.addAction(self.new_project_action)

        self.open_project_action = file_menu.addAction("打开配置")
        self.open_project_action.triggered.connect(self.open_project)
        toolbar.addAction(self.open_project_action)

        self.save_project_action = file_menu.addAction("保存配置")
        self.save_project_action.triggered.connect(self.save_project)
        toolbar.addAction(self.save_project_action)

        self.save_project_as_action = file_menu.addAction("配置另存为")
        self.save_project_as_action.triggered.connect(self.save_project_as)

        file_menu.addSeparator()
        self.open_files_action = file_menu.addAction("打开表格文件")
        self.open_files_action.triggered.connect(self.open_files)
        toolbar.addAction(self.open_files_action)

        self.save_current_file_action = file_menu.addAction("结果另存为文件")
        self.save_current_file_action.triggered.connect(self.save_current_file_as)

        file_menu.addSeparator()
        self.recent_projects_menu = QMenu("最近配置", self)
        file_menu.addMenu(self.recent_projects_menu)
        self.recent_files_menu = QMenu("最近文件", self)
        file_menu.addMenu(self.recent_files_menu)

        file_menu.addSeparator()
        exit_action = file_menu.addAction("退出")
        exit_action.triggered.connect(QApplication.quit)

        refresh_action = project_menu.addAction("重新加载当前文件")
        refresh_action.triggered.connect(self.reload_active_file)
        merge_action = project_menu.addAction("合并配置内文件")
        merge_action.triggered.connect(self.merge_project_files)

        split_action = tools_menu.addAction("按列拆分导出")
        split_action.triggered.connect(self.export_split_files)
        dedupe_action = tools_menu.addAction("切换去重")
        dedupe_action.triggered.connect(self._toggle_drop_duplicates)
        empty_action = tools_menu.addAction("切换移除空行")
        empty_action.triggered.connect(self._toggle_remove_empty_rows)

        about_action = help_menu.addAction("关于")
        about_action.triggered.connect(self.show_about)

    def _load_settings_into_form(self) -> None:
        geometry = self.settings.value("main_window_geometry")
        if geometry:
            self.restoreGeometry(geometry)

    def _save_settings(self) -> None:
        self.settings.setValue("main_window_geometry", self.saveGeometry())

    def closeEvent(self, event) -> None:  # type: ignore[override]
        self._save_settings()
        super().closeEvent(event)

    def new_project(self) -> None:
        self.project_config = ProjectConfig()
        self.project_path = ""
        self.current_original_df = pd.DataFrame()
        self.current_view_df = pd.DataFrame()
        self._push_config_to_form()
        self._rebuild_file_list()
        self._refresh_columns_table()
        self._update_preview(pd.DataFrame())
        self._log("已创建新配置")

    def open_project(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "打开配置", "", PROJECT_FILTER)
        if not path:
            return

        try:
            self.project_config = ProjectConfig.load(path)
        except Exception as exc:
            self._show_error(f"打开配置失败：{exc}")
            return

        self.project_path = path
        self._push_config_to_form()
        self._rebuild_file_list()
        self._add_recent_item("recent_projects", path)
        self._populate_recent_menus()
        self._log(f"已打开配置：{path}")

        if self.project_config.active_file:
            self._select_file_in_list(self.project_config.active_file)
            self.reload_active_file()
        else:
            self._update_preview(pd.DataFrame())

    def save_project(self) -> None:
        if not self.project_path:
            self.save_project_as()
            return
        self._sync_form_to_config()
        try:
            self.project_config.save(self.project_path)
            self._add_recent_item("recent_projects", self.project_path)
            self._populate_recent_menus()
            self._log(f"配置已保存：{self.project_path}")
        except Exception as exc:
            self._show_error(f"保存配置失败：{exc}")

    def save_project_as(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "配置另存为", "", PROJECT_FILTER)
        if not path:
            return
        if not path.endswith(".pypandas.json"):
            path = f"{path}.pypandas.json"
        self.project_path = path
        self.save_project()

    def open_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "打开表格文件", "", DATA_FILTER)
        if not files:
            return

        existing = list(self.project_config.source_files)
        for file_path in files:
            if file_path not in existing:
                existing.append(file_path)
            self._add_recent_item("recent_files", file_path)

        self.project_config.source_files = existing
        if not self.project_config.active_file:
            self.project_config.active_file = files[0]

        self._populate_recent_menus()
        self._rebuild_file_list()
        self._select_file_in_list(self.project_config.active_file)
        self.reload_active_file()

    def remove_selected_files(self) -> None:
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            return

        selected_paths = {
            item.data(Qt.ItemDataRole.UserRole) for item in selected_items
        }
        self.project_config.source_files = [
            path
            for path in self.project_config.source_files
            if path not in selected_paths
        ]
        if self.project_config.active_file in selected_paths:
            self.project_config.active_file = (
                self.project_config.source_files[0]
                if self.project_config.source_files
                else ""
            )

        self._rebuild_file_list()
        if self.project_config.active_file:
            self._select_file_in_list(self.project_config.active_file)
            self.reload_active_file()
        else:
            self.current_original_df = pd.DataFrame()
            self._refresh_columns_table()
            self._update_preview(pd.DataFrame())

    def choose_output_dir(self) -> None:
        directory = QFileDialog.getExistingDirectory(
            self, "选择输出目录", self.output_dir_edit.text()
        )
        if directory:
            self.output_dir_edit.setText(directory)

    def reload_active_file(self) -> None:
        self._sync_form_to_config()
        active_file = self.project_config.active_file
        if not active_file:
            self._show_error("请先打开一个表格文件")
            return

        try:
            dataframe = load_prepared_dataframe(
                active_file,
                header_row=self.project_config.header_row,
                sheet_name=self.project_config.sheet_name,
            )
            dataframe = clean_dataframe(
                dataframe,
                remove_empty_rows=self.project_config.remove_empty_rows,
                drop_duplicates=self.project_config.drop_duplicates,
                sort_column=self.project_config.sort_column,
            )
        except ExcelServiceError as exc:
            self._show_error(str(exc))
            return

        self.current_original_df = dataframe
        self.project_config.sync_columns([str(column) for column in dataframe.columns])
        self._refresh_columns_table()
        self._populate_split_combo()
        self._refresh_preview_from_current()
        self._log(f"已加载文件：{active_file}")

    def merge_project_files(self) -> None:
        self._sync_form_to_config()
        if not self.project_config.source_files:
            self._show_error("配置中没有可合并的文件")
            return

        try:
            merged = merge_files(
                self.project_config.source_files,
                header_row=self.project_config.header_row,
                sheet_name=self.project_config.sheet_name,
                remove_empty_rows=self.project_config.remove_empty_rows,
                drop_duplicates=self.project_config.drop_duplicates,
                sort_column=self.project_config.sort_column,
            )
        except ExcelServiceError as exc:
            self._show_error(str(exc))
            return

        self.current_original_df = merged
        self.project_config.sync_columns([str(column) for column in merged.columns])
        self._refresh_columns_table()
        self._populate_split_combo()
        self._refresh_preview_from_current()
        self._log(f"已合并 {len(self.project_config.source_files)} 个文件")

    def save_current_file_as(self) -> None:
        if self.current_view_df.empty and self.current_original_df.empty:
            self._show_error("当前没有可导出的数据")
            return

        path, selected_filter = QFileDialog.getSaveFileName(
            self,
            "导出当前结果",
            "",
            "Excel 文件 (*.xlsx);;CSV 文件 (*.csv);;Excel 启用宏文件 (*.xlsm)",
        )
        if not path:
            return
        path = self._ensure_export_suffix(path, selected_filter)

        self._refresh_preview_from_current()
        try:
            export_dataframe(self.current_view_df, path)
            self._log(f"已导出当前结果：{path}")
        except ExcelServiceError as exc:
            self._show_error(str(exc))

    def export_split_files(self) -> None:
        self._sync_form_to_config()
        if self.current_original_df.empty:
            self._show_error("当前没有可拆分的数据")
            return
        if not self.project_config.split_columns:
            self._show_error("请先选择拆分列")
            return
        if not self.project_config.output_dir:
            self._show_error("请先选择输出目录")
            return

        try:
            exported = split_and_export(
                self.current_original_df,
                split_columns=self.project_config.split_columns,
                output_dir=self.project_config.output_dir,
                output_prefix=self.project_config.output_prefix or "result",
                export_format=self.project_config.export_format,
                column_rules=self.project_config.column_rules,
                values=self.project_config.split_values,
            )
        except ExcelServiceError as exc:
            self._show_error(str(exc))
            return

        if not exported:
            self._show_error("没有匹配的拆分结果，请检查拆分值过滤")
            return

        self._log(
            f"已拆分导出 {len(exported)} 个文件到：{self.project_config.output_dir}"
        )
        QMessageBox.information(self, "导出完成", f"已生成 {len(exported)} 个文件")

    def show_about(self) -> None:
        QMessageBox.information(
            self,
            "关于 墨韩表格工具箱",
            "<h3>墨韩表格工具箱</h3>"
            "<p>支持配置保存、标题行识别、列重命名、列筛选、按列拆分导出、"
            "多文件合并、去重与空行清理。</p>"
            "<hr>"
            "<p>本软件使用了以下开源组件：</p>"
            "<p><b><a href='https://www.qt.io/'>Qt</a></b> &mdash; The Qt Company Ltd<br>"
            "Qt 是跨平台应用程序开发框架，采用 "
            "<a href='https://www.gnu.org/licenses/lgpl-3.0.html'>GNU LGPL v3</a> 授权。</p>"
            "<p><b>PySide6</b> &mdash; The Qt Company Ltd<br>"
            "Qt for Python (PySide6) 是 Qt 的官方 Python 绑定，采用 "
            "<a href='https://www.gnu.org/licenses/lgpl-3.0.html'>GNU LGPL v3</a> 授权。</p>",
        )

    def _handle_file_selection_changed(self) -> None:
        selected_items = self.file_list.selectedItems()
        if not selected_items:
            return
        selected_path = selected_items[0].data(Qt.ItemDataRole.UserRole)
        if selected_path != self.project_config.active_file:
            self.project_config.active_file = selected_path
            self.reload_active_file()

    def _handle_columns_table_changed(self, item: QTableWidgetItem) -> None:
        if self._updating_columns_table:
            return
        if item is None:
            return
        self._collect_column_rules_from_table()
        self._refresh_preview_from_current()

    def _handle_header_row_changed(self, value: int) -> None:
        self.project_config.header_row = value
        if self.project_config.active_file:
            self.reload_active_file()

    def _collect_column_rules_from_table(self) -> None:
        rules: list[ColumnRule] = []
        for row in range(self.columns_table.rowCount()):
            enabled_item = self.columns_table.item(row, 0)
            original_item = self.columns_table.item(row, 1)
            display_item = self.columns_table.item(row, 2)
            if enabled_item is None or original_item is None:
                continue
            rules.append(
                ColumnRule(
                    original_name=original_item.text(),
                    display_name=display_item.text()
                    if display_item
                    else original_item.text(),
                    enabled=enabled_item.checkState() == Qt.CheckState.Checked,
                )
            )
        self.project_config.column_rules = rules

    def _refresh_columns_table(self) -> None:
        self._updating_columns_table = True
        self.columns_table.blockSignals(True)
        self.columns_table.setRowCount(len(self.project_config.column_rules))
        for row, rule in enumerate(self.project_config.column_rules):
            enabled_item = QTableWidgetItem()
            enabled_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsUserCheckable
                | Qt.ItemFlag.ItemIsSelectable
            )
            enabled_item.setCheckState(
                Qt.CheckState.Checked if rule.enabled else Qt.CheckState.Unchecked
            )
            enabled_item.setBackground(QBrush(QColor("#f3f4f6")))
            self.columns_table.setItem(row, 0, enabled_item)

            original_item = QTableWidgetItem(rule.original_name)
            original_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
            )
            original_item.setBackground(QBrush(QColor("#f3f4f6")))
            self.columns_table.setItem(row, 1, original_item)

            display_item = QTableWidgetItem(rule.display_name or rule.original_name)
            display_item.setFlags(
                Qt.ItemFlag.ItemIsEnabled
                | Qt.ItemFlag.ItemIsSelectable
                | Qt.ItemFlag.ItemIsEditable
            )
            display_item.setBackground(QBrush(QColor("white")))
            display_item.setToolTip("用户输入区：在这里输入导出后的列名")
            if not rule.display_name or rule.display_name == rule.original_name:
                display_item.setForeground(QBrush(QColor("#7a7a7a")))
                hint_font = QFont(display_item.font())
                hint_font.setItalic(True)
                display_item.setFont(hint_font)
            else:
                active_font = QFont(display_item.font())
                active_font.setItalic(False)
                active_font.setWeight(QFont.Weight.Medium)
                display_item.setFont(active_font)
            self.columns_table.setItem(row, 2, display_item)

        self.columns_table.blockSignals(False)
        self._updating_columns_table = False

    def _refresh_preview_from_current(self) -> None:
        if self.current_original_df.empty:
            self._update_preview(pd.DataFrame())
            return

        self._collect_column_rules_from_table()
        transformed = apply_column_rules(
            self.current_original_df, self.project_config.column_rules
        )
        self._update_preview(transformed)

    def _update_preview(self, dataframe: pd.DataFrame) -> None:
        self.current_view_df = dataframe
        self.table_model.set_dataframe(dataframe)
        self.table_view.resizeColumnsToContents()
        if dataframe.empty:
            self.summary_label.setText("未加载数据")
        else:
            self.summary_label.setText(
                f"当前预览：{len(dataframe)} 行 / {len(dataframe.columns)} 列"
            )
        self._update_ui_state()

    def _update_ui_state(self) -> None:
        has_files = bool(self.project_config.source_files)
        has_data = not self.current_original_df.empty
        self.save_project_action.setEnabled(True)
        self.save_project_as_action.setEnabled(True)
        self.save_current_file_action.setEnabled(has_data)
        self.recent_files_menu.setEnabled(True)
        self.recent_projects_menu.setEnabled(True)
        self.status_bar.showMessage(
            f"配置文件 {len(self.project_config.source_files)} 个"
            if has_files
            else "未打开配置文件"
        )

    def _sync_form_to_config(self) -> None:
        self.project_config.sheet_name = self.sheet_name_edit.text().strip()
        self.project_config.header_row = self.header_row_spin.value()
        self.project_config.output_dir = self.output_dir_edit.text().strip()
        self.project_config.output_prefix = (
            self.output_prefix_edit.text().strip() or "result"
        )
        self.project_config.export_format = (
            self.export_format_combo.currentText().strip() or "xlsx"
        )
        self.project_config.split_columns = self.split_column_combo.checked_items()
        self.project_config.split_values = [
            line.strip()
            for line in self.split_values_edit.toPlainText().splitlines()
            if line.strip()
        ]
        self.project_config.sort_column = self.sort_column_edit.text().strip()
        self.project_config.remove_empty_rows = self.remove_empty_rows_check.isChecked()
        self.project_config.drop_duplicates = self.drop_duplicates_check.isChecked()
        self._collect_column_rules_from_table()

    def _push_config_to_form(self) -> None:
        self.sheet_name_edit.setText(self.project_config.sheet_name)
        self.header_row_spin.setValue(self.project_config.header_row)
        self.output_dir_edit.setText(self.project_config.output_dir)
        self.output_prefix_edit.setText(self.project_config.output_prefix)
        self.export_format_combo.setCurrentText(
            self.project_config.export_format or "xlsx"
        )
        self._populate_split_combo()
        self.split_values_edit.setPlainText("\n".join(self.project_config.split_values))
        self.sort_column_edit.setText(self.project_config.sort_column)
        self.remove_empty_rows_check.setChecked(self.project_config.remove_empty_rows)
        self.drop_duplicates_check.setChecked(self.project_config.drop_duplicates)
        self._refresh_columns_table()

    def _populate_split_combo(self) -> None:
        items: list[tuple[str, str]] = []
        for rule in self.project_config.column_rules:
            display = rule.display_name or rule.original_name
            items.append((display, rule.original_name))
        self.split_column_combo.populate_items(
            items, set(self.project_config.split_columns)
        )

    def _rebuild_file_list(self) -> None:
        self.file_list.clear()
        for path in self.project_config.source_files:
            item = QListWidgetItem(Path(path).name)
            item.setData(Qt.ItemDataRole.UserRole, path)
            item.setToolTip(path)
            self.file_list.addItem(item)
        self._update_ui_state()

    def _select_file_in_list(self, target_path: str) -> None:
        for row in range(self.file_list.count()):
            item = self.file_list.item(row)
            if item.data(Qt.ItemDataRole.UserRole) == target_path:
                self.file_list.setCurrentRow(row)
                break

    def _show_error(self, message: str) -> None:
        self._log(message)
        QMessageBox.critical(self, "错误", message)

    def _log(self, message: str) -> None:
        self.log_edit.appendPlainText(message)

    def _toggle_drop_duplicates(self) -> None:
        self.drop_duplicates_check.setChecked(
            not self.drop_duplicates_check.isChecked()
        )
        if not self.current_original_df.empty:
            self.reload_active_file()

    def _toggle_remove_empty_rows(self) -> None:
        self.remove_empty_rows_check.setChecked(
            not self.remove_empty_rows_check.isChecked()
        )
        if not self.current_original_df.empty:
            self.reload_active_file()

    def _ensure_export_suffix(self, path: str, selected_filter: str) -> str:
        if Path(path).suffix:
            return path
        if "*.csv" in selected_filter:
            return f"{path}.csv"
        if "*.xlsm" in selected_filter:
            return f"{path}.xlsm"
        return f"{path}.xlsx"

    def _add_recent_item(self, key: str, value: str, limit: int = 8) -> None:
        raw = self.settings.value(key, [])
        if not isinstance(raw, list):
            raw = []
        current: list[str] = [str(item) for item in raw if isinstance(item, str)]
        current = [item for item in current if item != value]
        current.insert(0, value)
        self.settings.setValue(key, current[:limit])

    def _populate_recent_menus(self) -> None:
        self.recent_projects_menu.clear()
        self.recent_files_menu.clear()

        raw_projects = self.settings.value("recent_projects", [])
        recent_projects: list[str] = (
            [str(item) for item in raw_projects if isinstance(item, str)]
            if isinstance(raw_projects, list)
            else []
        )
        raw_files = self.settings.value("recent_files", [])
        recent_files: list[str] = (
            [str(item) for item in raw_files if isinstance(item, str)]
            if isinstance(raw_files, list)
            else []
        )

        self._fill_recent_menu(
            self.recent_projects_menu, recent_projects, self._open_recent_project
        )
        self._fill_recent_menu(
            self.recent_files_menu, recent_files, self._open_recent_file
        )

    def _fill_recent_menu(self, menu: QMenu, values: list[str], callback) -> None:
        existing_values = [value for value in values if Path(value).exists()]
        if not existing_values:
            action = menu.addAction("暂无记录")
            action.setEnabled(False)
            return
        for value in existing_values:
            action = menu.addAction(value)
            action.triggered.connect(lambda checked=False, path=value: callback(path))

    def _open_recent_project(self, path: str) -> None:
        if not Path(path).exists():
            self._show_error(f"配置不存在：{path}")
            return
        try:
            self.project_config = ProjectConfig.load(path)
        except Exception as exc:
            self._show_error(f"打开配置失败：{exc}")
            return
        self.project_path = path
        self._push_config_to_form()
        self._rebuild_file_list()
        self._select_file_in_list(self.project_config.active_file)
        self.reload_active_file()

    def _open_recent_file(self, path: str) -> None:
        if not Path(path).exists():
            self._show_error(f"文件不存在：{path}")
            return
        if path not in self.project_config.source_files:
            self.project_config.source_files.append(path)
        self.project_config.active_file = path
        self._rebuild_file_list()
        self._select_file_in_list(path)
        self.reload_active_file()


def _wrap_layout(layout: QHBoxLayout) -> QWidget:
    widget = QWidget()
    widget.setLayout(layout)
    return widget
