# -*- coding: utf-8 -*-
"""
进销存数据访问层 (dal/inventory_dal.py)
涵盖：商品/库存、采购订单、销售订单 三张表的 CRUD
"""
from __future__ import annotations
import sqlite3
import logging
from typing import Optional
from database import db_manager

logger = logging.getLogger(__name__)


# ════════════════════════════════════════════════════════════════
#  商品 & 库存
# ════════════════════════════════════════════════════════════════

class ProductDAL:
    """商品/库存表数据访问对象"""

    def __init__(self):
        self._conn = db_manager.get_connection()

    def get_all(self, keyword: str = "") -> list[sqlite3.Row]:
        if keyword:
            sql = "SELECT * FROM products WHERE product_name LIKE ? ORDER BY product_id"
            return self._conn.execute(sql, (f"%{keyword}%",)).fetchall()
        return self._conn.execute(
            "SELECT * FROM products ORDER BY product_id"
        ).fetchall()

    def get_by_id(self, product_id: str) -> Optional[sqlite3.Row]:
        return self._conn.execute(
            "SELECT * FROM products WHERE product_id=?", (product_id,)
        ).fetchone()

    def get_max_id(self) -> str:
        row = self._conn.execute(
            "SELECT product_id FROM products ORDER BY product_id DESC LIMIT 1"
        ).fetchone()
        num = (int(row["product_id"][1:]) + 1) if row else 1
        return f"P{num:04d}"

    def insert(self, data: dict) -> bool:
        sql = """
            INSERT INTO products (product_id, product_name, unit, stock_qty, warning_qty)
            VALUES (:product_id, :product_name, :unit, :stock_qty, :warning_qty)
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("商品新增失败: %s", exc)
            raise

    def update_stock(self, product_id: str, delta: int) -> bool:
        """调整库存数量，delta 为正则增加，为负则减少"""
        try:
            self._conn.execute(
                "UPDATE products SET stock_qty = stock_qty + ? WHERE product_id=?",
                (delta, product_id),
            )
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("库存调整失败: %s", exc)
            raise

    def update(self, data: dict) -> bool:
        sql = """
            UPDATE products SET
                product_name=:product_name, unit=:unit,
                stock_qty=:stock_qty, warning_qty=:warning_qty
            WHERE product_id=:product_id
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("商品更新失败: %s", exc)
            raise

    def get_low_stock(self) -> list[sqlite3.Row]:
        """查询库存低于预警阈值的商品"""
        return self._conn.execute(
            "SELECT * FROM products WHERE stock_qty <= warning_qty ORDER BY stock_qty"
        ).fetchall()


# ════════════════════════════════════════════════════════════════
#  采购订单
# ════════════════════════════════════════════════════════════════

class PurchaseDAL:
    """采购订单表数据访问对象"""

    def __init__(self):
        self._conn = db_manager.get_connection()

    def get_all(
        self, keyword: str = "", page: int = 1, page_size: int = 20
    ) -> tuple[list[sqlite3.Row], int]:
        conditions = []
        params: list = []
        if keyword:
            conditions.append("(po.supplier LIKE ? OR p.product_name LIKE ?)")
            params.extend([f"%{keyword}%", f"%{keyword}%"])

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        total = self._conn.execute(
            f"SELECT COUNT(*) AS cnt FROM purchase_orders po "
            f"LEFT JOIN products p ON po.product_id=p.product_id {where}",
            params,
        ).fetchone()["cnt"]

        offset = (page - 1) * page_size
        sql = f"""
            SELECT po.*, p.product_name
            FROM purchase_orders po
            LEFT JOIN products p ON po.product_id=p.product_id
            {where}
            ORDER BY po.order_date DESC, po.po_id DESC
            LIMIT ? OFFSET ?
        """
        rows = self._conn.execute(sql, params + [page_size, offset]).fetchall()
        return rows, total

    def get_max_id(self) -> str:
        row = self._conn.execute(
            "SELECT po_id FROM purchase_orders ORDER BY po_id DESC LIMIT 1"
        ).fetchone()
        num = (int(row["po_id"][2:]) + 1) if row else 1
        return f"PO{num:06d}"

    def insert(self, data: dict) -> bool:
        sql = """
            INSERT INTO purchase_orders
                (po_id, supplier, product_id, quantity, unit_price, order_date, status, remark)
            VALUES
                (:po_id, :supplier, :product_id, :quantity, :unit_price,
                 :order_date, :status, :remark)
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("采购订单新增失败: %s", exc)
            raise


# ════════════════════════════════════════════════════════════════
#  销售订单
# ════════════════════════════════════════════════════════════════

class SalesDAL:
    """销售订单表数据访问对象"""

    def __init__(self):
        self._conn = db_manager.get_connection()

    def get_all(
        self, keyword: str = "", page: int = 1, page_size: int = 20
    ) -> tuple[list[sqlite3.Row], int]:
        conditions = []
        params: list = []
        if keyword:
            conditions.append("(so.customer LIKE ? OR p.product_name LIKE ?)")
            params.extend([f"%{keyword}%", f"%{keyword}%"])

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        total = self._conn.execute(
            f"SELECT COUNT(*) AS cnt FROM sales_orders so "
            f"LEFT JOIN products p ON so.product_id=p.product_id {where}",
            params,
        ).fetchone()["cnt"]

        offset = (page - 1) * page_size
        sql = f"""
            SELECT so.*, p.product_name
            FROM sales_orders so
            LEFT JOIN products p ON so.product_id=p.product_id
            {where}
            ORDER BY so.order_date DESC, so.so_id DESC
            LIMIT ? OFFSET ?
        """
        rows = self._conn.execute(sql, params + [page_size, offset]).fetchall()
        return rows, total

    def get_max_id(self) -> str:
        row = self._conn.execute(
            "SELECT so_id FROM sales_orders ORDER BY so_id DESC LIMIT 1"
        ).fetchone()
        num = (int(row["so_id"][2:]) + 1) if row else 1
        return f"SO{num:06d}"

    def insert(self, data: dict) -> bool:
        sql = """
            INSERT INTO sales_orders
                (so_id, customer, product_id, quantity, unit_price, order_date, status, remark)
            VALUES
                (:so_id, :customer, :product_id, :quantity, :unit_price,
                 :order_date, :status, :remark)
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("销售订单新增失败: %s", exc)
            raise
