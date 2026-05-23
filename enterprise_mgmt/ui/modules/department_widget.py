# -*- coding: utf-8 -*-
"""
部门与权限管理模块界面 (ui/modules/department_widget.py)
布局：左侧 QTreeWidget 展示层级架构，右侧显示部门详情及所属员工列表
"""
from __future__ import annotations
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter,
    QTreeWidget, QTreeWidgetItem, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QGroupBox, QFormLayout, QLineEdit,
    QTextEdit, QFrame, QMessageBox, QHeaderView, QDialog,
    QDialogButtonBox, QComboBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QFont, QIcon

from bll import DepartmentBLL, EmployeeBLL


class DepartmentDialog(QDialog):
    """部门新增 / 编辑对话框"""

    def __init__(self, parent=None, dept_data=None, all_depts=None):
        super().__init__(parent)
        self._dept_data = dept_data
        self._all_depts = all_depts or []
        self._is_edit   = dept_data is not None
        self.setWindowTitle("编辑部门" if self._is_edit else "新增部门")
        self.setMinimumWidth(380)
        self.setModal(True)
        self._build_ui()
        if self._is_edit:
            self._fill_form()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(10)

        title = QLabel("编辑部门" if self._is_edit else "新增部门")
        title.setObjectName("label_title")
        layout.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        self.edit_name    = QLineEdit()
        self.edit_manager = QLineEdit()
        self.combo_parent = QComboBox()
        self.combo_parent.addItem("— 无上级（顶层部门）—", None)
        for d in self._all_depts:
            self.combo_parent.addItem(f"{d['dept_name']}（{d['dept_id']}）", d["dept_id"])
        self.edit_desc = QTextEdit()
        self.edit_desc.setMaximumHeight(80)

        form.addRow("部门名称 *", self.edit_name)
        form.addRow("部门负责人",  self.edit_manager)
        form.addRow("上级部门",    self.combo_parent)
        form.addRow("职能描述",    self.edit_desc)
        layout.addLayout(form)

        btn_box = QDialogButtonBox()
        btn_save   = QPushButton("保  存")
        btn_cancel = QPushButton("取  消")
        btn_cancel.setObjectName("btn_secondary")
        btn_box.addButton(btn_save,   QDialogButtonBox.AcceptRole)
        btn_box.addButton(btn_cancel, QDialogButtonBox.RejectRole)
        btn_save.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_box)

    def _fill_form(self):
        d = self._dept_data
        self.edit_name.setText(d["dept_name"] or "")
        self.edit_manager.setText(d["manager"] or "")
        self.edit_desc.setPlainText(d["description"] or "")
        # 设置上级部门
        if d["parent_id"]:
            for i in range(self.combo_parent.count()):
                if self.combo_parent.itemData(i) == d["parent_id"]:
                    self.combo_parent.setCurrentIndex(i)
                    break

    def get_form_data(self) -> dict:
        return {
            "dept_name":   self.edit_name.text().strip(),
            "manager":     self.edit_manager.text().strip(),
            "parent_id":   self.combo_parent.currentData(),
            "description": self.edit_desc.toPlainText().strip(),
        }


class DepartmentWidget(QWidget):
    """部门管理主界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bll     = DepartmentBLL()
        self._emp_bll = EmployeeBLL()
        self._build_ui()
        self.load_tree()

    # ── UI 构建 ─────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(0)

        # ── 顶部操作栏 ───────────────────────────────────────────
        top_bar = QHBoxLayout()
        top_bar.setContentsMargins(0, 0, 0, 12)
        btn_add    = QPushButton("＋ 新增部门")
        btn_edit   = QPushButton("✏  编辑部门")
        btn_delete = QPushButton("🗑  删除部门")
        btn_add.setObjectName("btn_success")
        btn_edit.setObjectName("btn_secondary")
        btn_delete.setObjectName("btn_danger")
        btn_add.clicked.connect(self._on_add)
        btn_edit.clicked.connect(self._on_edit)
        btn_delete.clicked.connect(self._on_delete)
        top_bar.addWidget(btn_add)
        top_bar.addWidget(btn_edit)
        top_bar.addWidget(btn_delete)
        top_bar.addStretch()
        layout.addLayout(top_bar)

        # ── 主体分割区 ───────────────────────────────────────────
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(6)

        # 左：树形部门结构
        left_panel = QWidget()
        left_panel.setObjectName("card_widget")
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 8, 8, 8)

        lbl_tree = QLabel("  🏢  公司组织架构")
        lbl_tree.setObjectName("label_title")
        lbl_tree.setStyleSheet("font-size:14px; padding:6px 0;")
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setMinimumWidth(200)
        self.tree.currentItemChanged.connect(self._on_tree_select)

        left_layout.addWidget(lbl_tree)
        left_layout.addWidget(self.tree)

        # 右：部门详情 + 员工列表
        right_panel = QWidget()
        right_panel.setObjectName("card_widget")
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(16, 12, 16, 12)
        right_layout.setSpacing(12)

        # 部门详情卡
        detail_group = QGroupBox("部门详情")
        detail_form  = QFormLayout(detail_group)
        detail_form.setLabelAlignment(Qt.AlignRight)
        detail_form.setSpacing(8)

        self.lbl_dept_id   = QLineEdit(); self.lbl_dept_id.setReadOnly(True)
        self.lbl_dept_name = QLineEdit(); self.lbl_dept_name.setReadOnly(True)
        self.lbl_manager   = QLineEdit(); self.lbl_manager.setReadOnly(True)
        self.lbl_parent    = QLineEdit(); self.lbl_parent.setReadOnly(True)
        self.lbl_desc      = QTextEdit(); self.lbl_desc.setReadOnly(True)
        self.lbl_desc.setMaximumHeight(60)

        detail_form.addRow("部门编号:", self.lbl_dept_id)
        detail_form.addRow("部门名称:", self.lbl_dept_name)
        detail_form.addRow("负责人:",   self.lbl_manager)
        detail_form.addRow("上级部门:", self.lbl_parent)
        detail_form.addRow("职能描述:", self.lbl_desc)
        right_layout.addWidget(detail_group)

        # 所属员工表格
        emp_group  = QGroupBox("所属在职员工")
        emp_layout = QVBoxLayout(emp_group)
        self.emp_table = QTableWidget()
        self.emp_table.setColumnCount(4)
        self.emp_table.setHorizontalHeaderLabels(["工号", "姓名", "职位", "联系方式"])
        self.emp_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.emp_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.emp_table.setAlternatingRowColors(True)
        self.emp_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.emp_table.verticalHeader().setVisible(False)
        self.lbl_emp_count = QLabel("共 0 名在职员工")
        self.lbl_emp_count.setAlignment(Qt.AlignRight)

        emp_layout.addWidget(self.emp_table)
        emp_layout.addWidget(self.lbl_emp_count)
        right_layout.addWidget(emp_group)

        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([220, 580])

        layout.addWidget(splitter)

    # ── 数据加载 ─────────────────────────────────────────────────

    def load_tree(self):
        """重新构建部门树"""
        self.tree.clear()
        depts = self._bll.get_all()

        # 先建 id→item 映射，再连接父子关系
        id_item: dict[str, QTreeWidgetItem] = {}
        root_items: list[QTreeWidgetItem]   = []

        for d in depts:
            item = QTreeWidgetItem([f"  {d['dept_name']}  ({d['dept_id']})"])
            item.setData(0, Qt.UserRole, d["dept_id"])
            id_item[d["dept_id"]] = item

        for d in depts:
            item      = id_item[d["dept_id"]]
            parent_id = d["parent_id"]
            if parent_id and parent_id in id_item:
                id_item[parent_id].addChild(item)
            else:
                root_items.append(item)

        self.tree.addTopLevelItems(root_items)
        self.tree.expandAll()

    def _on_tree_select(self, current: QTreeWidgetItem, _):
        if current is None:
            return
        dept_id = current.data(0, Qt.UserRole)
        self._show_dept_detail(dept_id)
        self._show_dept_employees(dept_id)

    def _show_dept_detail(self, dept_id: str):
        dept = self._bll.get_dept(dept_id)
        if not dept:
            return
        self.lbl_dept_id.setText(dept["dept_id"])
        self.lbl_dept_name.setText(dept["dept_name"] or "")
        self.lbl_manager.setText(dept["manager"] or "（未设置）")
        # 获取上级部门名称
        parent_name = "无"
        if dept["parent_id"]:
            parent = self._bll.get_dept(dept["parent_id"])
            if parent:
                parent_name = parent["dept_name"]
        self.lbl_parent.setText(parent_name)
        self.lbl_desc.setPlainText(dept["description"] or "")

    def _show_dept_employees(self, dept_id: str):
        emps = self._emp_bll.get_employees_by_dept(dept_id)
        self.emp_table.setRowCount(len(emps))
        for r, emp in enumerate(emps):
            for c, val in enumerate([emp["emp_id"], emp["name"],
                                      emp["position"] or "", emp["phone"] or ""]):
                item = QTableWidgetItem(str(val))
                item.setTextAlignment(Qt.AlignCenter)
                self.emp_table.setItem(r, c, item)
        self.lbl_emp_count.setText(f"共 {len(emps)} 名在职员工")

    # ── 增删改 ──────────────────────────────────────────────────

    def _current_dept_id(self) -> str | None:
        item = self.tree.currentItem()
        return item.data(0, Qt.UserRole) if item else None

    def _on_add(self):
        all_depts = self._bll.get_all()
        dlg = DepartmentDialog(self, all_depts=all_depts)
        if dlg.exec_() != QDialog.Accepted:
            return
        form = dlg.get_form_data()
        try:
            dept_id = self._bll.add_department(form)
            self.load_tree()
            self._notify(f"部门【{form['dept_name']}】新增成功，编号：{dept_id}")
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "系统错误", str(exc))

    def _on_edit(self):
        dept_id = self._current_dept_id()
        if not dept_id:
            QMessageBox.warning(self, "提示", "请先在左侧选择要编辑的部门")
            return
        dept_row  = self._bll.get_dept(dept_id)
        all_depts = [d for d in self._bll.get_all() if d["dept_id"] != dept_id]
        dlg = DepartmentDialog(self, dept_data=dept_row, all_depts=all_depts)
        if dlg.exec_() != QDialog.Accepted:
            return
        form = dlg.get_form_data()
        form["dept_id"] = dept_id
        try:
            self._bll.update_department(form)
            self.load_tree()
            self._show_dept_detail(dept_id)
            self._notify(f"部门【{form['dept_name']}】信息已更新")
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "系统错误", str(exc))

    def _on_delete(self):
        dept_id = self._current_dept_id()
        if not dept_id:
            QMessageBox.warning(self, "提示", "请先选择要删除的部门")
            return
        dept = self._bll.get_dept(dept_id)
        ret  = QMessageBox.question(
            self, "确认删除",
            f"确定要删除部门【{dept['dept_name']}】吗？\n删除前请确保该部门无在职员工。",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if ret != QMessageBox.Yes:
            return
        try:
            self._bll.delete_department(dept_id)
            self.load_tree()
            self._clear_detail()
            self._notify(f"部门【{dept['dept_name']}】已删除")
        except ValueError as exc:
            QMessageBox.warning(self, "删除失败", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "系统错误", str(exc))

    def _clear_detail(self):
        for w in [self.lbl_dept_id, self.lbl_dept_name, self.lbl_manager, self.lbl_parent]:
            w.clear()
        self.lbl_desc.clear()
        self.emp_table.setRowCount(0)
        self.lbl_emp_count.setText("共 0 名在职员工")

    def _notify(self, msg: str):
        parent = self.window()
        if hasattr(parent, "statusBar"):
            parent.statusBar().showMessage(msg, 4000)
