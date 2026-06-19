from __future__ import annotations

import pandas as pd

from qtpy.QtCore import QAbstractTableModel, QModelIndex, Qt


class PandasModel(QAbstractTableModel):
    """Qt table model backed by a pandas DataFrame."""

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
