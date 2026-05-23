# -*- coding: utf-8 -*-
"""
人事管理模块界面 (ui/modules/hrm_widget.py)
功能：员工列表（分页）、搜索、新增、编辑、离职处理、导出Excel
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLineEdit, QComboBox, QLabel, QMessageBox,
    QHeaderView, QFileDialog, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QColor, QFont

from bll import EmployeeBLL, DepartmentBLL
from ui.dialogs import EmployeeDialog
from utils import export_to_excel
from config import PAGE_SIZE


class HrmWidget(QWidget):
    """人事管理模块主界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bll      = EmployeeBLL()
        self._dept_bll = DepartmentBLL()
        self._current_page = 1
        self._total_pages  = 1
        self._build_ui()
        self.load_data()

    # ── UI 构建 ─────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # ── 顶部搜索/操作栏 ──────────────────────────────────────
        top_bar = QHBoxLayout()

        self.edit_search = QLineEdit()
        self.edit_search.setPlaceholderText("🔍  搜索姓名 / 部门…")
        self.edit_search.setMaximumWidth(240)
        self.edit_search.returnPressed.connect(self._on_search)

        self.combo_dept = QComboBox()
        self.combo_dept.setMinimumWidth(140)
        self._fill_dept_combo()

        self.combo_status = QComboBox()
        self.combo_status.addItems(["在职", "离职", "全部"])

        btn_search = QPushButton("查  询")
        btn_search.clicked.connect(self._on_search)

        btn_add = QPushButton("＋ 新增员工")
        btn_add.setObjectName("btn_success")

        btn_resign = QPushButton("办理离职")
        btn_resign.setObjectName("btn_danger")

        btn_export = QPushButton("导出 Excel")
        btn_export.setObjectName("btn_secondary")

        btn_add.clicked.connect(self._on_add)
        btn_resign.clicked.connect(self._on_resign)
        btn_export.clicked.connect(self._on_export)

        top_bar.addWidget(self.edit_search)
        top_bar.addWidget(QLabel("部门:"))
        top_bar.addWidget(self.combo_dept)
        top_bar.addWidget(QLabel("状态:"))
        top_bar.addWidget(self.combo_status)
        top_bar.addWidget(btn_search)
        top_bar.addStretch()
        top_bar.addWidget(btn_add)
        top_bar.addWidget(btn_resign)
        top_bar.addWidget(btn_export)
        layout.addLayout(top_bar)

        # ── 表格 ─────────────────────────────────────────────────
        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "工号", "姓名", "性别", "部门", "职位",
            "基本工资(元)", "入职日期", "联系方式", "状态"
        ])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.verticalHeader().setVisible(False)
        self.table.doubleClicked.connect(self._on_edit)
        layout.addWidget(self.table)

        # ── 分页栏 ───────────────────────────────────────────────
        page_bar = QHBoxLayout()
        self.btn_prev = QPushButton("◀ 上一页")
        self.btn_next = QPushButton("下一页 ▶")
        self.btn_prev.setObjectName("btn_secondary")
        self.btn_next.setObjectName("btn_secondary")
        self.lbl_page = QLabel("第 1 页 / 共 1 页")
        self.lbl_page.setAlignment(Qt.AlignCenter)
        self.lbl_total = QLabel("共 0 条记录")

        self.btn_prev.clicked.connect(self._on_prev_page)
        self.btn_next.clicked.connect(self._on_next_page)

        page_bar.addStretch()
        page_bar.addWidget(self.lbl_total)
        page_bar.addSpacing(20)
        page_bar.addWidget(self.btn_prev)
        page_bar.addWidget(self.lbl_page)
        page_bar.addWidget(self.btn_next)
        layout.addLayout(page_bar)

    def _fill_dept_combo(self):
        self.combo_dept.clear()
        self.combo_dept.addItem("全部部门", "")
        self._dept_id_map: dict[str, str] = {"全部部门": ""}
        for d in self._dept_bll.get_all():
            self.combo_dept.addItem(d["dept_name"], d["dept_id"])

    # ── 数据加载 ─────────────────────────────────────────────────

    def load_data(self):
        keyword    = self.edit_search.text().strip()
        dept_id    = self.combo_dept.currentData() or ""
        status_txt = self.combo_status.currentText()
        status     = "" if status_txt == "全部" else status_txt

        rows, total, self._total_pages = self._bll.get_employee_list(
            keyword=keyword,
            dept_id=dept_id,
            status=status,
            page=self._current_page,
        )

        self.table.setRowCount(len(rows))
        for r_idx, row in enumerate(rows):
            values = [
                row["emp_id"], row["name"], row["gender"] or "",
                row["dept_name"] or "", row["position"] or "",
                f"{float(row['base_salary'] or 0):,.2f}",
                row["hire_date"] or "", row["phone"] or "", row["status"] or "",
            ]
            for c_idx, val in enumerate(values):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                # 离职员工置灰
                if row["status"] == "离职":
                    item.setForeground(QColor("#9E9E9E"))
                self.table.setItem(r_idx, c_idx, item)

        self.lbl_page.setText(f"第 {self._current_page} 页 / 共 {self._total_pages} 页")
        self.lbl_total.setText(f"共 {total} 条记录")
        self.btn_prev.setEnabled(self._current_page > 1)
        self.btn_next.setEnabled(self._current_page < self._total_pages)

    # ── 事件处理 ─────────────────────────────────────────────────

    def _on_search(self):
        self._current_page = 1
        self.load_data()

    def _on_prev_page(self):
        if self._current_page > 1:
            self._current_page -= 1
            self.load_data()

    def _on_next_page(self):
        if self._current_page < self._total_pages:
            self._current_page += 1
            self.load_data()

    def _on_add(self):
        dlg = EmployeeDialog(self)
        if dlg.exec_() == EmployeeDialog.Accepted:
            self.load_data()
            self._notify("员工入职登记成功")

    def _on_edit(self):
        row_idx = self.table.currentRow()
        if row_idx < 0:
            return
        emp_id  = self.table.item(row_idx, 0).text()
        emp_row = self._bll.get_employee(emp_id)
        if emp_row is None:
            return
        dlg = EmployeeDialog(self, emp_data=emp_row)
        if dlg.exec_() == EmployeeDialog.Accepted:
            self.load_data()
            self._notify("员工信息已更新")

    def _on_resign(self):
        row_idx = self.table.currentRow()
        if row_idx < 0:
            QMessageBox.warning(self, "提示", "请先选择要办理离职的员工")
            return
        emp_id   = self.table.item(row_idx, 0).text()
        emp_name = self.table.item(row_idx, 1).text()

        ret = QMessageBox.question(
            self, "确认离职",
            f"确定要为员工【{emp_name}（{emp_id}）】办理离职手续吗？\n此操作不可撤销。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ret == QMessageBox.Yes:
            try:
                self._bll.resign_employee(emp_id)
                self.load_data()
                self._notify(f"员工 {emp_name} 已办理离职")
            except ValueError as exc:
                QMessageBox.warning(self, "操作失败", str(exc))

    def _on_export(self):
        filepath, _ = QFileDialog.getSaveFileName(
            self, "导出 Excel", "员工信息.xlsx",
            "Excel 文件 (*.xlsx)"
        )
        if not filepath:
            return
        try:
            # 导出当前查询的全量数据（不分页）
            keyword = self.edit_search.text().strip()
            dept_id = self.combo_dept.currentData() or ""
            status_txt = self.combo_status.currentText()
            status = "" if status_txt == "全部" else status_txt

            all_rows, _, _ = self._bll.get_employee_list(
                keyword=keyword, dept_id=dept_id, status=status,
                page=1
            )
            # 获取全量（简单做法：直接改 page_size；这里复用 BLL）
            from dal import EmployeeDAL
            raw_rows, _ = EmployeeDAL().get_all(
                keyword=keyword, dept_id=dept_id, status=status,
                page=1, page_size=100_000
            )
            data = [
                [r["emp_id"], r["name"], r["gender"] or "",
                 r["dept_name"] or "", r["position"] or "",
                 float(r["base_salary"] or 0),
                 r["hire_date"] or "", r["phone"] or "", r["status"] or ""]
                for r in raw_rows
            ]
            headers = ["工号", "姓名", "性别", "部门", "职位", "基本工资", "入职日期", "联系方式", "状态"]
            export_to_excel(filepath, headers, data, sheet_name="员工信息", title="员工信息导出")
            QMessageBox.information(self, "导出成功", f"文件已保存到：\n{filepath}")
        except Exception as exc:
            QMessageBox.critical(self, "导出失败", str(exc))

    def _notify(self, msg: str):
        """通知主窗口状态栏"""
        parent = self.window()
        if hasattr(parent, "statusBar"):
            parent.statusBar().showMessage(msg, 4000)
