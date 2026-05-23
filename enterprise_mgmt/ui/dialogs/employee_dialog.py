# -*- coding: utf-8 -*-
"""
员工新增/编辑对话框 (ui/dialogs/employee_dialog.py)
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox,
    QPushButton, QDialogButtonBox, QLabel, QMessageBox
)
from PyQt5.QtCore import Qt, QDate
from bll import EmployeeBLL, DepartmentBLL


class EmployeeDialog(QDialog):
    """员工新增 / 编辑对话框"""

    def __init__(self, parent=None, emp_data=None):
        """
        :param emp_data: sqlite3.Row 对象；为 None 则为新增模式
        """
        super().__init__(parent)
        self._bll      = EmployeeBLL()
        self._dept_bll = DepartmentBLL()
        self._emp_data = emp_data
        self._is_edit  = emp_data is not None

        self.setWindowTitle("编辑员工信息" if self._is_edit else "新增员工")
        self.setMinimumWidth(420)
        self.setModal(True)
        self._build_ui()
        if self._is_edit:
            self._fill_form()

    # ── UI 构建 ─────────────────────────────────────────────────

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 16)
        main_layout.setSpacing(12)

        # 标题
        title = QLabel("编辑员工信息" if self._is_edit else "新增员工")
        title.setObjectName("label_title")
        main_layout.addWidget(title)

        # 表单
        form = QFormLayout()
        form.setRowWrapPolicy(QFormLayout.WrapLongRows)
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        self.edit_name    = QLineEdit()
        self.combo_gender = QComboBox()
        self.combo_gender.addItems(["男", "女"])
        self.combo_dept   = QComboBox()
        self._load_departments()
        self.edit_position = QLineEdit()
        self.spin_salary   = QDoubleSpinBox()
        self.spin_salary.setRange(0, 9_999_999)
        self.spin_salary.setDecimals(2)
        self.spin_salary.setSuffix("  元")
        self.date_hire     = QDateEdit(QDate.currentDate())
        self.date_hire.setCalendarPopup(True)
        self.date_hire.setDisplayFormat("yyyy-MM-dd")
        self.edit_phone    = QLineEdit()
        self.combo_status  = QComboBox()
        self.combo_status.addItems(["在职", "离职"])

        form.addRow("姓    名 *", self.edit_name)
        form.addRow("性    别", self.combo_gender)
        form.addRow("所属部门 *", self.combo_dept)
        form.addRow("职    位", self.edit_position)
        form.addRow("基本工资 *", self.spin_salary)
        form.addRow("入职日期 *", self.date_hire)
        form.addRow("联系方式", self.edit_phone)
        form.addRow("员工状态", self.combo_status)

        main_layout.addLayout(form)

        # 按钮
        btn_box = QDialogButtonBox()
        self.btn_save   = QPushButton("保  存")
        self.btn_cancel = QPushButton("取  消")
        self.btn_cancel.setObjectName("btn_secondary")
        btn_box.addButton(self.btn_save,   QDialogButtonBox.AcceptRole)
        btn_box.addButton(self.btn_cancel, QDialogButtonBox.RejectRole)
        self.btn_save.clicked.connect(self._on_save)
        self.btn_cancel.clicked.connect(self.reject)
        main_layout.addWidget(btn_box)

    def _load_departments(self):
        """填充部门下拉框"""
        self.combo_dept.clear()
        self._dept_id_map: dict[str, str] = {}  # dept_name -> dept_id
        for dept in self._dept_bll.get_all():
            self.combo_dept.addItem(dept["dept_name"])
            self._dept_id_map[dept["dept_name"]] = dept["dept_id"]

    # ── 填充已有数据 ─────────────────────────────────────────────

    def _fill_form(self):
        d = self._emp_data
        self.edit_name.setText(d["name"] or "")
        idx = self.combo_gender.findText(d["gender"] or "男")
        self.combo_gender.setCurrentIndex(max(0, idx))

        # 设置部门
        dept_row = self._dept_bll.get_dept(d["dept_id"])
        if dept_row:
            idx = self.combo_dept.findText(dept_row["dept_name"])
            self.combo_dept.setCurrentIndex(max(0, idx))

        self.edit_position.setText(d["position"] or "")
        self.spin_salary.setValue(float(d["base_salary"] or 0))
        if d["hire_date"]:
            self.date_hire.setDate(QDate.fromString(d["hire_date"], "yyyy-MM-dd"))
        self.edit_phone.setText(d["phone"] or "")
        idx = self.combo_status.findText(d["status"] or "在职")
        self.combo_status.setCurrentIndex(max(0, idx))

    # ── 保存逻辑 ─────────────────────────────────────────────────

    def _on_save(self):
        dept_name = self.combo_dept.currentText()
        dept_id   = self._dept_id_map.get(dept_name, "")
        form = {
            "name":        self.edit_name.text().strip(),
            "gender":      self.combo_gender.currentText(),
            "dept_id":     dept_id,
            "position":    self.edit_position.text().strip(),
            "base_salary": self.spin_salary.value(),
            "hire_date":   self.date_hire.date().toString("yyyy-MM-dd"),
            "phone":       self.edit_phone.text().strip(),
            "status":      self.combo_status.currentText(),
        }
        if self._is_edit:
            form["emp_id"] = self._emp_data["emp_id"]

        try:
            if self._is_edit:
                self._bll.update_employee(form)
                QMessageBox.information(self, "成功", f"员工 {form['name']} 信息已更新")
            else:
                emp_id = self._bll.add_employee(form)
                QMessageBox.information(self, "成功", f"员工 {form['name']} 入职登记完成，工号：{emp_id}")
            self.accept()
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "系统错误", f"操作失败：{exc}")
