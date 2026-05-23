# -*- coding: utf-8 -*-
"""
程序入口 (main.py)
职责：
  1. 配置日志
  2. 初始化数据库（首次运行建表）
  3. 启动 PyQt5 应用并显示主窗口
"""

import sys
import logging
import os

# ── 将项目根目录加入 Python 路径（支持直接运行）──────────────────
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# ── 日志配置 ────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  [%(levelname)s]  %(name)s: %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("app.log", encoding="utf-8"),
    ],
)
logger = logging.getLogger("main")


def main():
    # ── PyQt5 高 DPI 支持 ────────────────────────────────────────
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import Qt

    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps,   True)

    app = QApplication(sys.argv)
    app.setApplicationName("企业内部业务管理系统")
    app.setOrganizationName("Enterprise Corp")

    # ── 数据库初始化 ─────────────────────────────────────────────
    try:
        from database import db_manager
        db_manager.initialize()
        logger.info("数据库初始化完成")
    except Exception as exc:
        from PyQt5.QtWidgets import QMessageBox
        QMessageBox.critical(
            None, "数据库初始化失败",
            f"无法初始化数据库，程序将退出。\n\n错误信息：{exc}"
        )
        sys.exit(1)

    # ── 创建并显示主窗口 ─────────────────────────────────────────
    from ui.main_window import MainWindow
    window = MainWindow()
    window.show()

    logger.info("应用启动成功")
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
