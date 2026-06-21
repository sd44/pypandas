from __future__ import annotations

import os
from pathlib import Path
import sys

from qtpy.QtGui import QIcon
from qtpy.QtCore import QTranslator, QLocale, QLibraryInfo
from qtpy.QtWidgets import QApplication

from mohanxlsx.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("墨韩表格工具箱")
    app.setOrganizationName("墨韩")

    icon_path = Path(__file__).parent / "mohanxlsx.svg"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    translator = QTranslator()
    # 获取 Qt 安装的 translations 目录路径 (兼容 PySide2/Qt5 和 PySide6/Qt6)
    try:
        translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    except AttributeError:
        translations_path = QLibraryInfo.location(QLibraryInfo.TranslationsPath)
    translator.load(QLocale(), "qtbase", "_", translations_path)
    app.installTranslator(translator)

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
