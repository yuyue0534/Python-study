# -*- coding: utf-8 -*-
"""
进销存业务逻辑层 (bll/inventory_bll.py)
核心规则：
  1. 采购入库 → 增加库存 + 记录支出流水
  2. 销售出库 → 减少库存（校验库存充足性）+ 记录收入流水
  3. 销售/采购后自动触发库存预警检查
"""

import logging
from datetime import date
from dal import ProductDAL, PurchaseDAL, SalesDAL, FinanceDAL
from config import PAGE_SIZE

logger = logging.getLogger(__name__)


class InventoryBLL:
    """进销存业务逻辑服务"""

    def __init__(self):
        self._product_dal  = ProductDAL()
        self._purchase_dal = PurchaseDAL()
        self._sales_dal    = SalesDAL()
        self._fin_dal      = FinanceDAL()

    # ── 商品管理 ────────────────────────────────────────────────

    def get_products(self, keyword: str = ""):
        return self._product_dal.get_all(keyword)

    def add_product(self, form: dict) -> str:
        self._validate_product(form)
        pid = self._product_dal.get_max_id()
        form["product_id"] = pid
        self._product_dal.insert(form)
        return pid

    def update_product(self, form: dict) -> None:
        if not form.get("product_id"):
            raise ValueError("商品编号不能为空")
        self._validate_product(form)
        self._product_dal.update(form)

    def get_low_stock(self):
        return self._product_dal.get_low_stock()

    # ── 采购入库 ────────────────────────────────────────────────

    def create_purchase(self, form: dict) -> str:
        """
        创建采购订单 + 更新库存 + 记录财务支出
        返回新采购单号
        """
        # 基础校验
        if not form.get("supplier", "").strip():
            raise ValueError("供应商不能为空")
        product_id = form.get("product_id")
        if not product_id or not self._product_dal.get_by_id(product_id):
            raise ValueError("请选择有效商品")

        qty = self._parse_int(form.get("quantity"), "采购数量")
        price = self._parse_float(form.get("unit_price"), "采购单价")

        form["quantity"]   = qty
        form["unit_price"] = price
        form.setdefault("order_date", str(date.today()))
        form.setdefault("status", "已入库")
        form.setdefault("remark", "")

        po_id = self._purchase_dal.get_max_id()
        form["po_id"] = po_id

        # 写订单 → 增库存 → 写财务（三步要保持一致性）
        self._purchase_dal.insert(form)
        self._product_dal.update_stock(product_id, qty)
        self._fin_dal.insert({
            "record_type": "支出",
            "amount":       qty * price,
            "category":    "采购支出",
            "dept_id":     None,
            "ref_id":      po_id,
            "record_date": form["order_date"],
            "description": f"采购入库 {po_id} ({form['supplier']})",
        })
        logger.info("采购入库完成: %s, 商品=%s, 数量=%d", po_id, product_id, qty)
        return po_id

    def get_purchase_list(self, keyword: str = "", page: int = 1):
        rows, total = self._purchase_dal.get_all(keyword, page, PAGE_SIZE)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        return list(rows), total, total_pages

    # ── 销售出库 ────────────────────────────────────────────────

    def create_sale(self, form: dict) -> str:
        """
        创建销售订单 + 扣减库存 + 记录财务收入
        返回新销售单号
        """
        if not form.get("customer", "").strip():
            raise ValueError("客户名称不能为空")
        product_id = form.get("product_id")
        if not product_id:
            raise ValueError("请选择商品")

        product = self._product_dal.get_by_id(product_id)
        if not product:
            raise ValueError("商品不存在")

        qty   = self._parse_int(form.get("quantity"), "销售数量")
        price = self._parse_float(form.get("unit_price"), "销售单价")

        # 库存充足性校验
        if product["stock_qty"] < qty:
            raise ValueError(
                f"库存不足！当前库存 {product['stock_qty']} {product['unit']}，"
                f"本次销售 {qty} {product['unit']}"
            )

        form["quantity"]   = qty
        form["unit_price"] = price
        form.setdefault("order_date", str(date.today()))
        form.setdefault("status", "已完成")
        form.setdefault("remark", "")

        so_id = self._sales_dal.get_max_id()
        form["so_id"] = so_id

        self._sales_dal.insert(form)
        self._product_dal.update_stock(product_id, -qty)   # 扣减
        self._fin_dal.insert({
            "record_type": "收入",
            "amount":       qty * price,
            "category":    "销售收入",
            "dept_id":     None,
            "ref_id":      so_id,
            "record_date": form["order_date"],
            "description": f"销售出库 {so_id} ({form['customer']})",
        })
        logger.info("销售出库完成: %s, 商品=%s, 数量=%d", so_id, product_id, qty)
        return so_id

    def get_sales_list(self, keyword: str = "", page: int = 1):
        rows, total = self._sales_dal.get_all(keyword, page, PAGE_SIZE)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        return list(rows), total, total_pages

    # ── 工具方法 ────────────────────────────────────────────────

    @staticmethod
    def _parse_int(val, field: str) -> int:
        try:
            v = int(val)
            if v <= 0:
                raise ValueError
            return v
        except (TypeError, ValueError):
            raise ValueError(f"{field} 必须为正整数")

    @staticmethod
    def _parse_float(val, field: str) -> float:
        try:
            v = float(val)
            if v < 0:
                raise ValueError
            return v
        except (TypeError, ValueError):
            raise ValueError(f"{field} 必须为非负数字")

    @staticmethod
    def _validate_product(form: dict) -> None:
        if not (form.get("product_name") or "").strip():
            raise ValueError("商品名称不能为空")
        try:
            qty = int(form.get("stock_qty", 0))
            form["stock_qty"] = qty
        except (TypeError, ValueError):
            raise ValueError("库存数量必须为整数")
        try:
            wq = int(form.get("warning_qty", 10))
            form["warning_qty"] = wq
        except (TypeError, ValueError):
            raise ValueError("预警阈值必须为整数")
