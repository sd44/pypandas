from __future__ import annotations

from pathlib import Path
import sys

from qtpy.QtGui import QIcon
from qtpy.QtCore import QTranslator, QLibraryInfo
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
    # 显式加载中文翻译，避免依赖运行环境 locale 导致标准右键菜单仍为英文。
    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    translator.load("qtbase_zh_CN", translations_path)
    app.installTranslator(translator)

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
