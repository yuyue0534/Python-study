# -*- coding: utf-8 -*-
"""
主窗口 (ui/main_window.py)
布局：左侧深色导航栏 + 顶部工具栏 + 右侧 QStackedWidget 工作区
"""

from datetime import datetime
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QStackedWidget, QPushButton, QLabel, QStatusBar,
    QSizePolicy, QFrame, QMessageBox, QAction, QToolBar
)
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtGui import QFont, QIcon

from ui.styles import MAIN_STYLE
from ui.modules import HrmWidget, DepartmentWidget, PsiWidget, FinanceWidget
from config import APP_NAME, APP_VERSION


class MainWindow(QMainWindow):
    """企业管理系统主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v{APP_VERSION}")
        self.setMinimumSize(1280, 820)
        self.resize(1440, 900)

        # 应用全局 QSS
        self.setStyleSheet(MAIN_STYLE)

        self._build_ui()
        self._start_clock()

        # 默认打开第一个模块
        self._switch_module(0)

    # ── UI 构建 ─────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QHBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # ── 左侧导航面板 ──────────────────────────────────────────
        nav_panel = QWidget()
        nav_panel.setObjectName("nav_panel")
        nav_layout = QVBoxLayout(nav_panel)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(0)

        # LOGO / 系统名称
        logo_label = QLabel(f"  🏢  {APP_NAME}")
        logo_label.setObjectName("nav_logo_label")
        logo_label.setWordWrap(True)
        logo_label.setMinimumHeight(70)
        nav_layout.addWidget(logo_label)

        # 分隔线
        line = QFrame(); line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color: #3D5166;")
        nav_layout.addWidget(line)

        # 导航按钮（使用 QButtonGroup 实现单选高亮）
        self._nav_buttons: list[QPushButton] = []
        nav_items = [
            ("👥  员工人事管理",   0),
            ("🏢  部门组织架构",   1),
            ("📦  进销存管理",     2),
            ("📊  财务报表中心",   3),
        ]
        for text, idx in nav_items:
            btn = QPushButton(text)
            btn.setCheckable(True)
            btn.setMinimumHeight(50)
            btn.clicked.connect(lambda checked, i=idx: self._switch_module(i))
            nav_layout.addWidget(btn)
            self._nav_buttons.append(btn)

        nav_layout.addStretch()

        # 版本信息
        ver_label = QLabel(f"  v{APP_VERSION}")
        ver_label.setStyleSheet("color: #546E7A; font-size: 11px; padding: 8px 16px;")
        nav_layout.addWidget(ver_label)

        root_layout.addWidget(nav_panel)

        # ── 右侧主体区 ───────────────────────────────────────────
        right_area = QVBoxLayout()
        right_area.setContentsMargins(0, 0, 0, 0)
        right_area.setSpacing(0)

        # 顶部工具栏
        top_bar = self._build_top_bar()
        right_area.addWidget(top_bar)

        # 分隔线
        sep = QFrame(); sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #E0E0E0;")
        right_area.addWidget(sep)

        # 工作区（StackedWidget，每个模块一页）
        self._stack = QStackedWidget()
        self._hrm_widget  = HrmWidget()
        self._dept_widget = DepartmentWidget()
        self._psi_widget  = PsiWidget()
        self._fin_widget  = FinanceWidget()
        self._stack.addWidget(self._hrm_widget)
        self._stack.addWidget(self._dept_widget)
        self._stack.addWidget(self._psi_widget)
        self._stack.addWidget(self._fin_widget)
        right_area.addWidget(self._stack)

        right_wrapper = QWidget()
        right_wrapper.setLayout(right_area)
        root_layout.addWidget(right_wrapper)

        # ── 状态栏 ───────────────────────────────────────────────
        self._status_bar = QStatusBar()
        self._status_bar.showMessage("就绪")
        self.setStatusBar(self._status_bar)

        # 状态栏右侧时间
        self._lbl_time = QLabel()
        self._lbl_time.setStyleSheet("color: #FFFFFF; padding-right: 12px;")
        self._status_bar.addPermanentWidget(self._lbl_time)

    def _build_top_bar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("top_bar")
        bar.setFixedHeight(50)

        layout = QHBoxLayout(bar)
        layout.setContentsMargins(16, 0, 16, 0)

        self._lbl_module_title = QLabel("员工人事管理")
        layout.addWidget(self._lbl_module_title)
        layout.addStretch()

        # 快捷操作按钮
        btn_refresh = QPushButton("🔄  刷新")
        btn_about   = QPushButton("ℹ  关于")
        btn_exit    = QPushButton("✕  退出")
        btn_exit.setObjectName("btn_danger")

        btn_refresh.clicked.connect(self._on_refresh)
        btn_about.clicked.connect(self._on_about)
        btn_exit.clicked.connect(self.close)

        for btn in [btn_refresh, btn_about, btn_exit]:
            layout.addWidget(btn)

        return bar

    # ── 导航切换 ─────────────────────────────────────────────────

    _MODULE_TITLES = ["员工人事管理", "部门组织架构", "进销存管理", "财务报表中心"]

    def _switch_module(self, index: int):
        # 更新按钮高亮
        for i, btn in enumerate(self._nav_buttons):
            btn.setChecked(i == index)

        self._stack.setCurrentIndex(index)
        self._lbl_module_title.setText(self._MODULE_TITLES[index])
        self._status_bar.showMessage(f"已切换至：{self._MODULE_TITLES[index]}", 3000)

    # ── 顶栏操作 ─────────────────────────────────────────────────

    def _on_refresh(self):
        idx = self._stack.currentIndex()
        widget = self._stack.currentWidget()
        # 调用各模块的刷新方法
        if hasattr(widget, "load_data"):
            widget.load_data()
        elif hasattr(widget, "load_all"):
            widget.load_all()
        elif hasattr(widget, "load_tree"):
            widget.load_tree()
        elif hasattr(widget, "refresh_dashboard"):
            widget.refresh_dashboard()
        self._status_bar.showMessage("数据已刷新", 2000)

    def _on_about(self):
        QMessageBox.about(
            self,
            f"关于 {APP_NAME}",
            f"<h3>{APP_NAME}</h3>"
            f"<p>版本：<b>v{APP_VERSION}</b></p>"
            f"<p>技术栈：Python 3.10+ / PyQt5 / SQLite</p>"
            f"<p>架构：DAL → BLL → UI 三层分层架构</p>"
            f"<hr>"
            f"<p style='color:#607D8B'>本系统仅供内部使用</p>"
        )

    # ── 时钟 ─────────────────────────────────────────────────────

    def _start_clock(self):
        timer = QTimer(self)
        timer.timeout.connect(self._update_clock)
        timer.start(1000)
        self._update_clock()

    def _update_clock(self):
        now = datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
        self._lbl_time.setText(now)

    # ── 关闭事件 ─────────────────────────────────────────────────

    def closeEvent(self, event):
        ret = QMessageBox.question(
            self, "确认退出",
            "确定要退出系统吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            from database import db_manager
            db_manager.close()
            event.accept()
        else:
            event.ignore()
