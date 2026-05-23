# -*- coding: utf-8 -*-
"""
进销存业务模块界面 (ui/modules/psi_widget.py)
三个子 Tab：库存管理 | 采购管理 | 销售管理
库存低于阈值时行高亮警告
"""

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QTableWidget, QTableWidgetItem, QPushButton,
    QLineEdit, QLabel, QHeaderView, QMessageBox,
    QDialog, QFormLayout, QSpinBox, QDoubleSpinBox,
    QDialogButtonBox, QComboBox, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

from bll import InventoryBLL
from ui.dialogs.psi_dialogs import PurchaseDialog, SalesDialog
from config import PAGE_SIZE


# ════════════════════════════════════════════════════════════════
#  商品新增 / 编辑对话框
# ════════════════════════════════════════════════════════════════

class ProductDialog(QDialog):
    def __init__(self, parent=None, product_data=None):
        super().__init__(parent)
        self._data    = product_data
        self._is_edit = product_data is not None
        self.setWindowTitle("编辑商品" if self._is_edit else "新增商品")
        self.setMinimumWidth(360)
        self.setModal(True)
        self._build_ui()
        if self._is_edit:
            self._fill()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 16)

        title = QLabel("编辑商品" if self._is_edit else "新增商品")
        title.setObjectName("label_title")
        layout.addWidget(title)

        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.setSpacing(10)

        self.edit_name     = QLineEdit()
        self.edit_unit     = QLineEdit(); self.edit_unit.setPlaceholderText("件/箱/吨…")
        self.spin_stock    = QSpinBox();  self.spin_stock.setRange(0, 9_999_999)
        self.spin_warning  = QSpinBox();  self.spin_warning.setRange(0, 9_999_999)
        self.spin_warning.setValue(10)

        form.addRow("商品名称 *", self.edit_name)
        form.addRow("计量单位",   self.edit_unit)
        form.addRow("初始库存",   self.spin_stock)
        form.addRow("预警阈值",   self.spin_warning)
        layout.addLayout(form)

        btn_box    = QDialogButtonBox()
        btn_save   = QPushButton("保  存")
        btn_cancel = QPushButton("取  消")
        btn_cancel.setObjectName("btn_secondary")
        btn_box.addButton(btn_save,   QDialogButtonBox.AcceptRole)
        btn_box.addButton(btn_cancel, QDialogButtonBox.RejectRole)
        btn_save.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_box)

    def _fill(self):
        d = self._data
        self.edit_name.setText(d["product_name"] or "")
        self.edit_unit.setText(d["unit"] or "")
        self.spin_stock.setValue(int(d["stock_qty"] or 0))
        self.spin_warning.setValue(int(d["warning_qty"] or 10))

    def get_form_data(self) -> dict:
        return {
            "product_name": self.edit_name.text().strip(),
            "unit":         self.edit_unit.text().strip() or "件",
            "stock_qty":    self.spin_stock.value(),
            "warning_qty":  self.spin_warning.value(),
        }


# ════════════════════════════════════════════════════════════════
#  进销存主界面
# ════════════════════════════════════════════════════════════════

class PsiWidget(QWidget):
    """进销存模块主界面（含三个子标签页）"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bll = InventoryBLL()
        self._purchase_page = 1
        self._sales_page    = 1
        self._build_ui()
        self.load_all()

    # ── UI 构建 ─────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_inventory_tab(), "📦  库存管理")
        self.tabs.addTab(self._build_purchase_tab(),  "🛒  采购管理")
        self.tabs.addTab(self._build_sales_tab(),     "💰  销售管理")
        layout.addWidget(self.tabs)

    # ── 库存 Tab ────────────────────────────────────────────────

    def _build_inventory_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        # 操作栏
        bar = QHBoxLayout()
        self.inv_search = QLineEdit()
        self.inv_search.setPlaceholderText("🔍 搜索商品名称…")
        self.inv_search.setMaximumWidth(220)
        self.inv_search.returnPressed.connect(self.load_inventory)

        btn_search  = QPushButton("查询")
        btn_add_p   = QPushButton("＋ 新增商品")
        btn_edit_p  = QPushButton("✏ 编辑商品")
        btn_add_p.setObjectName("btn_success")
        btn_edit_p.setObjectName("btn_secondary")
        btn_search.clicked.connect(self.load_inventory)
        btn_add_p.clicked.connect(self._on_add_product)
        btn_edit_p.clicked.connect(self._on_edit_product)

        self.lbl_warning_count = QLabel()
        self.lbl_warning_count.setObjectName("label_warning")

        bar.addWidget(self.inv_search)
        bar.addWidget(btn_search)
        bar.addStretch()
        bar.addWidget(self.lbl_warning_count)
        bar.addWidget(btn_add_p)
        bar.addWidget(btn_edit_p)
        layout.addLayout(bar)

        # 库存表格
        self.inv_table = QTableWidget()
        self.inv_table.setObjectName("inventory_table")
        self.inv_table.setColumnCount(6)
        self.inv_table.setHorizontalHeaderLabels(
            ["商品编号", "商品名称", "计量单位", "当前库存", "预警阈值", "库存状态"]
        )
        self.inv_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.inv_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.inv_table.setAlternatingRowColors(True)
        self.inv_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.inv_table.verticalHeader().setVisible(False)
        layout.addWidget(self.inv_table)
        return w

    # ── 采购 Tab ────────────────────────────────────────────────

    def _build_purchase_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        bar = QHBoxLayout()
        self.po_search = QLineEdit()
        self.po_search.setPlaceholderText("🔍 搜索供应商 / 商品…")
        self.po_search.setMaximumWidth(220)
        self.po_search.returnPressed.connect(lambda: self._po_search_go())
        btn_search = QPushButton("查询")
        btn_new    = QPushButton("＋ 新增采购")
        btn_new.setObjectName("btn_success")
        btn_search.clicked.connect(lambda: self._po_search_go())
        btn_new.clicked.connect(self._on_new_purchase)

        bar.addWidget(self.po_search)
        bar.addWidget(btn_search)
        bar.addStretch()
        bar.addWidget(btn_new)
        layout.addLayout(bar)

        self.po_table = QTableWidget()
        self.po_table.setColumnCount(7)
        self.po_table.setHorizontalHeaderLabels(
            ["采购单号", "供应商", "商品名称", "数量", "单价(元)", "总金额(元)", "采购日期"]
        )
        self.po_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.po_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.po_table.setAlternatingRowColors(True)
        self.po_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.po_table.verticalHeader().setVisible(False)
        layout.addWidget(self.po_table)

        # 分页
        page_bar = QHBoxLayout()
        self.po_btn_prev = QPushButton("◀ 上一页"); self.po_btn_prev.setObjectName("btn_secondary")
        self.po_btn_next = QPushButton("下一页 ▶"); self.po_btn_next.setObjectName("btn_secondary")
        self.po_lbl_page = QLabel("第 1 页")
        self.po_btn_prev.clicked.connect(self._po_prev)
        self.po_btn_next.clicked.connect(self._po_next)
        page_bar.addStretch()
        page_bar.addWidget(self.po_btn_prev)
        page_bar.addWidget(self.po_lbl_page)
        page_bar.addWidget(self.po_btn_next)
        layout.addLayout(page_bar)
        return w

    # ── 销售 Tab ────────────────────────────────────────────────

    def _build_sales_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        bar = QHBoxLayout()
        self.so_search = QLineEdit()
        self.so_search.setPlaceholderText("🔍 搜索客户 / 商品…")
        self.so_search.setMaximumWidth(220)
        self.so_search.returnPressed.connect(lambda: self._so_search_go())
        btn_search = QPushButton("查询")
        btn_new    = QPushButton("＋ 新增销售")
        btn_new.setObjectName("btn_success")
        btn_search.clicked.connect(lambda: self._so_search_go())
        btn_new.clicked.connect(self._on_new_sale)

        bar.addWidget(self.so_search)
        bar.addWidget(btn_search)
        bar.addStretch()
        bar.addWidget(btn_new)
        layout.addLayout(bar)

        self.so_table = QTableWidget()
        self.so_table.setColumnCount(7)
        self.so_table.setHorizontalHeaderLabels(
            ["销售单号", "客户名称", "商品名称", "数量", "单价(元)", "总金额(元)", "销售日期"]
        )
        self.so_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.so_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.so_table.setAlternatingRowColors(True)
        self.so_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.so_table.verticalHeader().setVisible(False)
        layout.addWidget(self.so_table)

        page_bar = QHBoxLayout()
        self.so_btn_prev = QPushButton("◀ 上一页"); self.so_btn_prev.setObjectName("btn_secondary")
        self.so_btn_next = QPushButton("下一页 ▶"); self.so_btn_next.setObjectName("btn_secondary")
        self.so_lbl_page = QLabel("第 1 页")
        self.so_btn_prev.clicked.connect(self._so_prev)
        self.so_btn_next.clicked.connect(self._so_next)
        page_bar.addStretch()
        page_bar.addWidget(self.so_btn_prev)
        page_bar.addWidget(self.so_lbl_page)
        page_bar.addWidget(self.so_btn_next)
        layout.addLayout(page_bar)
        return w

    # ── 数据加载 ─────────────────────────────────────────────────

    def load_all(self):
        self.load_inventory()
        self.load_purchase()
        self.load_sales()

    def load_inventory(self):
        keyword = self.inv_search.text().strip()
        products = self._bll.get_products(keyword)
        low_stock = self._bll.get_low_stock()
        low_ids   = {p["product_id"] for p in low_stock}

        self.inv_table.setRowCount(len(products))
        for r, p in enumerate(products):
            is_low    = p["product_id"] in low_ids
            status    = "⚠ 库存不足" if is_low else "✓ 正常"
            values    = [
                p["product_id"], p["product_name"], p["unit"] or "件",
                str(p["stock_qty"]), str(p["warning_qty"]), status
            ]
            for c, val in enumerate(values):
                item = QTableWidgetItem(val)
                item.setTextAlignment(Qt.AlignCenter)
                if is_low:
                    item.setBackground(QColor("#FFEBEE"))
                    item.setForeground(QColor("#C62828"))
                self.inv_table.setItem(r, c, item)

        warn_cnt = len(low_ids)
        if warn_cnt:
            self.lbl_warning_count.setText(f"⚠ {warn_cnt} 件商品库存不足！")
        else:
            self.lbl_warning_count.setText("")

    def load_purchase(self):
        kw = self.po_search.text().strip()
        rows, total, total_pages = self._bll.get_purchase_list(kw, self._purchase_page)
        self._po_total_pages = total_pages
        self.po_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            vals = [
                row["po_id"], row["supplier"], row["product_name"] or "",
                str(row["quantity"]),
                f"{float(row['unit_price']):,.2f}",
                f"{float(row['total_amount']):,.2f}",
                row["order_date"] or "",
            ]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setTextAlignment(Qt.AlignCenter)
                self.po_table.setItem(r, c, item)
        self.po_lbl_page.setText(f"第 {self._purchase_page} / {total_pages} 页  共 {total} 条")
        self.po_btn_prev.setEnabled(self._purchase_page > 1)
        self.po_btn_next.setEnabled(self._purchase_page < total_pages)

    def load_sales(self):
        kw = self.so_search.text().strip()
        rows, total, total_pages = self._bll.get_sales_list(kw, self._sales_page)
        self._so_total_pages = total_pages
        self.so_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            vals = [
                row["so_id"], row["customer"], row["product_name"] or "",
                str(row["quantity"]),
                f"{float(row['unit_price']):,.2f}",
                f"{float(row['total_amount']):,.2f}",
                row["order_date"] or "",
            ]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setTextAlignment(Qt.AlignCenter)
                self.so_table.setItem(r, c, item)
        self.so_lbl_page.setText(f"第 {self._sales_page} / {total_pages} 页  共 {total} 条")
        self.so_btn_prev.setEnabled(self._sales_page > 1)
        self.so_btn_next.setEnabled(self._sales_page < total_pages)

    # ── 分页事件 ─────────────────────────────────────────────────

    def _po_search_go(self):
        self._purchase_page = 1; self.load_purchase()

    def _po_prev(self):
        if self._purchase_page > 1:
            self._purchase_page -= 1; self.load_purchase()

    def _po_next(self):
        if self._purchase_page < getattr(self, "_po_total_pages", 1):
            self._purchase_page += 1; self.load_purchase()

    def _so_search_go(self):
        self._sales_page = 1; self.load_sales()

    def _so_prev(self):
        if self._sales_page > 1:
            self._sales_page -= 1; self.load_sales()

    def _so_next(self):
        if self._sales_page < getattr(self, "_so_total_pages", 1):
            self._sales_page += 1; self.load_sales()

    # ── 商品增删改 ──────────────────────────────────────────────

    def _on_add_product(self):
        dlg = ProductDialog(self)
        if dlg.exec_() != QDialog.Accepted:
            return
        try:
            pid = self._bll.add_product(dlg.get_form_data())
            self.load_inventory()
            self._notify(f"商品新增成功，编号：{pid}")
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))

    def _on_edit_product(self):
        row = self.inv_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "提示", "请先选择要编辑的商品")
            return
        pid = self.inv_table.item(row, 0).text()
        products = self._bll.get_products()
        p_data = next((p for p in products if p["product_id"] == pid), None)
        if not p_data:
            return
        dlg = ProductDialog(self, product_data=p_data)
        if dlg.exec_() != QDialog.Accepted:
            return
        form = dlg.get_form_data()
        form["product_id"] = pid
        try:
            self._bll.update_product(form)
            self.load_inventory()
            self._notify("商品信息已更新")
        except ValueError as exc:
            QMessageBox.warning(self, "校验失败", str(exc))

    # ── 采购 / 销售 ──────────────────────────────────────────────

    def _on_new_purchase(self):
        dlg = PurchaseDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            self.load_inventory()
            self.load_purchase()
            self._notify("采购入库成功")

    def _on_new_sale(self):
        dlg = SalesDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            self.load_inventory()
            self.load_sales()
            self._notify("销售出库成功")

    def _notify(self, msg: str):
        p = self.window()
        if hasattr(p, "statusBar"):
            p.statusBar().showMessage(msg, 4000)
