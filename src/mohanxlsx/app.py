from __future__ import annotations

import os
from pathlib import Path
import sys

from qtpy.QtGui import QIcon


def _setup_input_method() -> None:
    """在导入 Qt 之前配置输入法支持。

    - Linux: 将系统 Qt6 输入法插件目录加入 QT_PLUGIN_PATH，
      以便 PySide6 捆绑的 Qt 能够发现 fcitx5 / ibus 等输入法插件。
    - Windows: 不设置 QT_IM_MODULE，让 Qt 使用系统默认的
      IME (Input Method Editor) 支持中文/日文/韩文等输入法。
    """
    if sys.platform == "win32":
        # Windows 下不干预，让 Qt 使用原生 IME 支持
        # 显式清除可能被设置为 Linux 专用值的环境变量
        os.environ.pop("QT_IM_MODULE", None)
        return

    # --- Linux 输入法配置 ---
    # 默认 IM 模块：fcitx 兼容 fcitx4/5，ibus 兼容 ibus/fcitx5-ibus
    if "QT_IM_MODULE" not in os.environ:
        os.environ["QT_IM_MODULE"] = "fcitx"

    # 系统 Qt6 插件目录（包含 libfcitx5platforminputcontextplugin.so 等）
    system_qt6_plugins = "/usr/lib/x86_64-linux-gnu/qt6/plugins"
    existing = os.environ.get("QT_PLUGIN_PATH", "")
    if os.path.isdir(system_qt6_plugins):
        if existing:
            os.environ["QT_PLUGIN_PATH"] = system_qt6_plugins + os.pathsep + existing
        else:
            os.environ["QT_PLUGIN_PATH"] = system_qt6_plugins


_setup_input_method()

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
    # 获取 Qt 安装的 translations 目录路径
    translations_path = QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath)
    translator.load(QLocale(), "qtbase", "_", translations_path)
    app.installTranslator(translator)

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
