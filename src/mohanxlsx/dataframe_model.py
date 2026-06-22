from __future__ import annotations

import pandas as pd

from qtpy.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt


class PandasModel(QAbstractTableModel):
    """Qt table model backed by a pandas DataFrame."""

    RawValueRole = int(Qt.ItemDataRole.UserRole) + 1

    def __init__(self, dataframe: pd.DataFrame | None = None, parent=None) -> None:
        super().__init__(parent)
        self._dataframe = dataframe if dataframe is not None else pd.DataFrame()

    def set_dataframe(self, dataframe: pd.DataFrame | None) -> None:
        self.beginResetModel()
        self._dataframe = dataframe if dataframe is not None else pd.DataFrame()
        self.endResetModel()

    def dataframe(self) -> pd.DataFrame:
        return self._dataframe

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._dataframe.index)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._dataframe.columns)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        value = self._dataframe.iat[index.row(), index.column()]
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            if pd.isna(value):
                return ""
            return str(value)

        if role == self.RawValueRole:
            if pd.isna(value):
                return None
            return value

        if role == Qt.ItemDataRole.TextAlignmentRole:
            return int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        return None

    def headerData(
        self,
        section: int,
        orientation: Qt.Orientation,
        role: int = Qt.ItemDataRole.DisplayRole,
    ):
        if role != Qt.ItemDataRole.DisplayRole:
            return None

        if orientation == Qt.Orientation.Horizontal:
            if 0 <= section < len(self._dataframe.columns):
                return str(self._dataframe.columns[section])
            return ""

        if 0 <= section < len(self._dataframe.index):
            return str(section + 1)
        return ""


class DataFrameSortProxyModel(QSortFilterProxyModel):
    """Sort DataFrame rows by raw values instead of display strings."""

    def lessThan(self, left: QModelIndex, right: QModelIndex) -> bool:
        source_model = self.sourceModel()
        if source_model is None:
            return super().lessThan(left, right)

        left_value = source_model.data(left, PandasModel.RawValueRole)
        right_value = source_model.data(right, PandasModel.RawValueRole)
        return self._sort_key(left_value) < self._sort_key(right_value)

    def _sort_key(self, value):
        if value is None:
            return (3, "")
        if isinstance(value, bool):
            return (0, int(value))
        if isinstance(value, (int, float)):
            return (0, value)
        if isinstance(value, pd.Timestamp):
            return (1, value.to_pydatetime())
        if isinstance(value, str):
            return (2, value.casefold())
        return (2, str(value).casefold())
