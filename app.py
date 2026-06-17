from __future__ import annotations

import sys

from qtpy.QtCore import QTranslator, QLocale, QLibraryInfo
from qtpy.QtWidgets import QApplication

from main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("墨韩表格工具箱")
    app.setOrganizationName("墨韩")

    translator = QTranslator()
    # 获取 Qt 安装的 translations 目录路径
    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    translator.load(QLocale(), "qtbase", "_", translations_path)
    app.installTranslator(translator)

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
