# -*- coding: utf-8 -*-
"""
QSS 样式表 (ui/styles.py)
定义全局商务扁平化主题：深色侧边栏 + 浅色工作区 + 扁平按钮
"""

MAIN_STYLE = """
/* ════════════════════════════════
   全局字体与背景
════════════════════════════════ */
QWidget {
    font-family: "Microsoft YaHei", "PingFang SC", "Segoe UI", sans-serif;
    font-size: 13px;
    color: #2C3E50;
    background-color: #F5F7FA;
}

/* ════════════════════════════════
   主窗口
════════════════════════════════ */
QMainWindow {
    background-color: #F5F7FA;
}

/* ════════════════════════════════
   左侧导航面板
════════════════════════════════ */
#nav_panel {
    background-color: #2E4057;
    min-width: 200px;
    max-width: 200px;
}

#nav_logo_label {
    color: #FFFFFF;
    font-size: 15px;
    font-weight: bold;
    padding: 20px 16px 10px 16px;
    background-color: #1E2D3F;
}

#nav_panel QPushButton {
    text-align: left;
    padding: 12px 20px;
    border: none;
    border-radius: 0px;
    color: #B0BEC5;
    background-color: transparent;
    font-size: 13px;
}
#nav_panel QPushButton:hover {
    background-color: #3D5166;
    color: #ECEFF1;
}
#nav_panel QPushButton:checked {
    background-color: #1565C0;
    color: #FFFFFF;
    font-weight: bold;
    border-left: 4px solid #64B5F6;
    padding-left: 16px;
}

/* ════════════════════════════════
   顶部工具栏
════════════════════════════════ */
#top_bar {
    background-color: #FFFFFF;
    border-bottom: 1px solid #E0E0E0;
    min-height: 50px;
    max-height: 50px;
}
#top_bar QLabel {
    font-size: 17px;
    font-weight: bold;
    color: #1565C0;
    padding-left: 10px;
}
#top_bar QPushButton {
    border: none;
    border-radius: 4px;
    padding: 6px 14px;
    color: #546E7A;
    background-color: transparent;
}
#top_bar QPushButton:hover {
    background-color: #E3F2FD;
    color: #1565C0;
}

/* ════════════════════════════════
   TabWidget 工作区
════════════════════════════════ */
QTabWidget::pane {
    border: none;
    background-color: #F5F7FA;
}
QTabBar::tab {
    background-color: #ECEFF1;
    color: #607D8B;
    padding: 8px 20px;
    border: 1px solid #CFD8DC;
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    margin-right: 2px;
}
QTabBar::tab:selected {
    background-color: #FFFFFF;
    color: #1565C0;
    font-weight: bold;
    border-top: 2px solid #1565C0;
}
QTabBar::tab:hover:!selected {
    background-color: #E3F2FD;
    color: #1565C0;
}

/* ════════════════════════════════
   通用卡片容器
════════════════════════════════ */
#card_widget {
    background-color: #FFFFFF;
    border: 1px solid #E8EAF0;
    border-radius: 8px;
}

/* ════════════════════════════════
   表格
════════════════════════════════ */
QTableWidget, QTableView {
    background-color: #FFFFFF;
    gridline-color: #ECEFF1;
    border: 1px solid #E0E0E0;
    border-radius: 4px;
    selection-background-color: #BBDEFB;
    selection-color: #0D47A1;
    alternate-background-color: #F5F8FF;
}
QTableWidget::item, QTableView::item {
    padding: 6px 10px;
    border: none;
}
QHeaderView::section {
    background-color: #37474F;
    color: #ECEFF1;
    font-weight: bold;
    padding: 8px 10px;
    border: none;
    border-right: 1px solid #546E7A;
}
QHeaderView::section:last {
    border-right: none;
}

/* ════════════════════════════════
   按钮
════════════════════════════════ */
QPushButton {
    background-color: #1565C0;
    color: #FFFFFF;
    border: none;
    border-radius: 4px;
    padding: 7px 18px;
    font-size: 13px;
}
QPushButton:hover {
    background-color: #1976D2;
}
QPushButton:pressed {
    background-color: #0D47A1;
}
QPushButton:disabled {
    background-color: #B0BEC5;
    color: #ECEFF1;
}

QPushButton#btn_secondary {
    background-color: #ECEFF1;
    color: #37474F;
    border: 1px solid #CFD8DC;
}
QPushButton#btn_secondary:hover {
    background-color: #E3F2FD;
    color: #1565C0;
    border-color: #90CAF9;
}

QPushButton#btn_danger {
    background-color: #C62828;
    color: #FFFFFF;
}
QPushButton#btn_danger:hover {
    background-color: #D32F2F;
}

QPushButton#btn_success {
    background-color: #2E7D32;
    color: #FFFFFF;
}
QPushButton#btn_success:hover {
    background-color: #388E3C;
}

QPushButton#btn_warning {
    background-color: #E65100;
    color: #FFFFFF;
}
QPushButton#btn_warning:hover {
    background-color: #EF6C00;
}

/* ════════════════════════════════
   输入框
════════════════════════════════ */
QLineEdit, QTextEdit, QPlainTextEdit {
    background-color: #FFFFFF;
    border: 1px solid #CFD8DC;
    border-radius: 4px;
    padding: 6px 10px;
    color: #2C3E50;
    selection-background-color: #BBDEFB;
}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus {
    border: 1px solid #1565C0;
    background-color: #FAFCFF;
}
QLineEdit:read-only {
    background-color: #ECEFF1;
    color: #607D8B;
}

QComboBox {
    background-color: #FFFFFF;
    border: 1px solid #CFD8DC;
    border-radius: 4px;
    padding: 6px 10px;
    min-width: 100px;
}
QComboBox:focus {
    border: 1px solid #1565C0;
}
QComboBox::drop-down {
    border: none;
    width: 20px;
}
QComboBox QAbstractItemView {
    background-color: #FFFFFF;
    border: 1px solid #CFD8DC;
    selection-background-color: #E3F2FD;
    selection-color: #1565C0;
}

QDateEdit {
    background-color: #FFFFFF;
    border: 1px solid #CFD8DC;
    border-radius: 4px;
    padding: 6px 10px;
}
QDateEdit:focus {
    border: 1px solid #1565C0;
}

/* ════════════════════════════════
   SpinBox
════════════════════════════════ */
QSpinBox, QDoubleSpinBox {
    background-color: #FFFFFF;
    border: 1px solid #CFD8DC;
    border-radius: 4px;
    padding: 6px 10px;
}
QSpinBox:focus, QDoubleSpinBox:focus {
    border: 1px solid #1565C0;
}

/* ════════════════════════════════
   树形视图
════════════════════════════════ */
QTreeView, QTreeWidget {
    background-color: #FFFFFF;
    border: 1px solid #E0E0E0;
    border-radius: 4px;
    outline: none;
}
QTreeView::item, QTreeWidget::item {
    height: 28px;
    padding-left: 6px;
}
QTreeView::item:selected, QTreeWidget::item:selected {
    background-color: #BBDEFB;
    color: #0D47A1;
}
QTreeView::item:hover, QTreeWidget::item:hover {
    background-color: #E3F2FD;
}
QTreeView::branch:has-children:!has-siblings:closed,
QTreeView::branch:closed:has-children:has-siblings {
    image: url(none);
}

/* ════════════════════════════════
   滚动条
════════════════════════════════ */
QScrollBar:vertical {
    background-color: #F5F7FA;
    width: 8px;
    margin: 0;
}
QScrollBar::handle:vertical {
    background-color: #B0BEC5;
    border-radius: 4px;
    min-height: 20px;
}
QScrollBar::handle:vertical:hover {
    background-color: #78909C;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0;
}
QScrollBar:horizontal {
    background-color: #F5F7FA;
    height: 8px;
}
QScrollBar::handle:horizontal {
    background-color: #B0BEC5;
    border-radius: 4px;
    min-width: 20px;
}
QScrollBar::handle:horizontal:hover {
    background-color: #78909C;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0;
}

/* ════════════════════════════════
   状态栏
════════════════════════════════ */
QStatusBar {
    background-color: #1565C0;
    color: #FFFFFF;
    font-size: 12px;
    padding: 2px 10px;
}
QStatusBar::item {
    border: none;
}

/* ════════════════════════════════
   分组框
════════════════════════════════ */
QGroupBox {
    font-weight: bold;
    border: 1px solid #CFD8DC;
    border-radius: 6px;
    margin-top: 10px;
    padding-top: 8px;
    background-color: #FFFFFF;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: #1565C0;
}

/* ════════════════════════════════
   对话框
════════════════════════════════ */
QDialog {
    background-color: #F5F7FA;
}

/* ════════════════════════════════
   库存预警高亮行（通过代码设置背景色）
════════════════════════════════ */
QTableWidget#inventory_table QTableWidgetItem[warning="true"] {
    background-color: #FFEBEE;
    color: #C62828;
}

/* ════════════════════════════════
   标签
════════════════════════════════ */
QLabel#label_title {
    font-size: 16px;
    font-weight: bold;
    color: #1565C0;
}
QLabel#label_stat {
    font-size: 20px;
    font-weight: bold;
    color: #1565C0;
}
QLabel#label_warning {
    color: #C62828;
    font-weight: bold;
}

/* ════════════════════════════════
   分隔线
════════════════════════════════ */
QFrame[frameShape="4"],  /* HLine */
QFrame[frameShape="5"] { /* VLine */
    color: #E0E0E0;
}
"""
