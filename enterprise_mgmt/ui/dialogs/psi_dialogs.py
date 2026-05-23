# -*- coding: utf-8 -*-
"""
采购录入对话框 & 销售录入对话框 (ui/dialogs/psi_dialogs.py)
"""

from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QFormLayout,
    QLineEdit, QComboBox, QDateEdit, QSpinBox, QDoubleSpinBox,
    QPushButton, QDialogButtonBox, QLabel, QMessageBox, QTextEdit
)
from PyQt5.QtCore import Qt, QDate
from bll import InventoryBLL


# ════════════════════════════════════════════════════════════════
#  采购录入对话框
# ════════════════════════════════════════════════════════════════

class PurchaseDialog(QDialog):
    """采购订单录入对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bll = InventoryBLL()
        self.setWindowTitle("新增采购订单")
        self.setMinimumWidth(400)
        self.setModal(True)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(10)

        title = QLabel("新增采购订单")
        title.setObjectName("label_title")
        layout.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        self.edit_supplier = QLineEdit()
        self.combo_product = QComboBox()
        self._load_products()
        self.spin_qty      = QSpinBox()
        self.spin_qty.setRange(1, 999_999)
        self.spin_price    = QDoubleSpinBox()
        self.spin_price.setRange(0, 9_999_999)
        self.spin_price.setDecimals(2)
        self.spin_price.setSuffix("  元")
        self.date_order    = QDateEdit(QDate.currentDate())
        self.date_order.setCalendarPopup(True)
        self.date_order.setDisplayFormat("yyyy-MM-dd")
        self.edit_remark   = QLineEdit()

        form.addRow("供 应 商 *", self.edit_supplier)
        form.addRow("商    品 *", self.combo_product)
        form.addRow("采购数量 *", self.spin_qty)
        form.addRow("采购单价 *", self.spin_price)
        form.addRow("采购日期",   self.date_order)
        form.addRow("备    注",   self.edit_remark)
        layout.addLayout(form)

        btn_box = QDialogButtonBox()
        self.btn_save   = QPushButton("确认入库")
        self.btn_cancel = QPushButton("取  消")
        self.btn_cancel.setObjectName("btn_secondary")
        btn_box.addButton(self.btn_save,   QDialogButtonBox.AcceptRole)
        btn_box.addButton(self.btn_cancel, QDialogButtonBox.RejectRole)
        self.btn_save.clicked.connect(self._on_save)
        self.btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_box)

    def _load_products(self):
        self._product_id_map: dict[str, str] = {}
        for p in self._bll.get_products():
            label = f"{p['product_name']} ({p['product_id']})"
            self.combo_product.addItem(label)
            self._product_id_map[label] = p["product_id"]

    def _on_save(self):
        label      = self.combo_product.currentText()
        product_id = self._product_id_map.get(label, "")
        form = {
            "supplier":   self.edit_supplier.text().strip(),
            "product_id": product_id,
            "quantity":   self.spin_qty.value(),
            "unit_price": self.spin_price.value(),
            "order_date": self.date_order.date().toString("yyyy-MM-dd"),
            "remark":     self.edit_remark.text().strip(),
        }
        try:
            po_id = self._bll.create_purchase(form)
            QMessageBox.information(self, "采购入库成功", f"采购单号：{po_id}")
            self.accept()
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "系统错误", f"操作失败：{exc}")


# ════════════════════════════════════════════════════════════════
#  销售录入对话框
# ════════════════════════════════════════════════════════════════

class SalesDialog(QDialog):
    """销售订单录入对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bll = InventoryBLL()
        self.setWindowTitle("新增销售订单")
        self.setMinimumWidth(400)
        self.setModal(True)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)
        layout.setSpacing(10)

        title = QLabel("新增销售订单")
        title.setObjectName("label_title")
        layout.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        self.edit_customer = QLineEdit()
        self.combo_product = QComboBox()
        self._load_products()
        self.spin_qty    = QSpinBox()
        self.spin_qty.setRange(1, 999_999)
        self.spin_price  = QDoubleSpinBox()
        self.spin_price.setRange(0, 9_999_999)
        self.spin_price.setDecimals(2)
        self.spin_price.setSuffix("  元")
        self.date_order  = QDateEdit(QDate.currentDate())
        self.date_order.setCalendarPopup(True)
        self.date_order.setDisplayFormat("yyyy-MM-dd")
        self.edit_remark = QLineEdit()

        form.addRow("客户名称 *", self.edit_customer)
        form.addRow("商    品 *", self.combo_product)
        form.addRow("销售数量 *", self.spin_qty)
        form.addRow("成交单价 *", self.spin_price)
        form.addRow("销售日期",   self.date_order)
        form.addRow("备    注",   self.edit_remark)
        layout.addLayout(form)

        btn_box = QDialogButtonBox()
        self.btn_save   = QPushButton("确认销售")
        self.btn_cancel = QPushButton("取  消")
        self.btn_cancel.setObjectName("btn_secondary")
        btn_box.addButton(self.btn_save,   QDialogButtonBox.AcceptRole)
        btn_box.addButton(self.btn_cancel, QDialogButtonBox.RejectRole)
        self.btn_save.clicked.connect(self._on_save)
        self.btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_box)

    def _load_products(self):
        self._product_id_map: dict[str, str] = {}
        for p in self._bll.get_products():
            label = f"{p['product_name']}  [库存:{p['stock_qty']}{p['unit']}]"
            self.combo_product.addItem(label)
            self._product_id_map[label] = p["product_id"]

    def _on_save(self):
        label      = self.combo_product.currentText()
        product_id = self._product_id_map.get(label, "")
        form = {
            "customer":   self.edit_customer.text().strip(),
            "product_id": product_id,
            "quantity":   self.spin_qty.value(),
            "unit_price": self.spin_price.value(),
            "order_date": self.date_order.date().toString("yyyy-MM-dd"),
            "remark":     self.edit_remark.text().strip(),
        }
        try:
            so_id = self._bll.create_sale(form)
            QMessageBox.information(self, "销售成功", f"销售单号：{so_id}")
            self.accept()
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))
        except Exception as exc:
            QMessageBox.critical(self, "系统错误", f"操作失败：{exc}")
