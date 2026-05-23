# -*- coding: utf-8 -*-
"""
财务数据访问层 (dal/finance_dal.py)
提供财务流水查询、月度统计、部门支出汇总等接口
"""
from __future__ import annotations
import sqlite3
import logging
from database import db_manager

logger = logging.getLogger(__name__)


class FinanceDAL:
    """财务流水表数据访问对象"""

    def __init__(self):
        self._conn = db_manager.get_connection()

    def insert(self, data: dict) -> bool:
        """记录一条财务流水"""
        sql = """
            INSERT INTO finance_records
                (record_type, amount, category, dept_id, ref_id, record_date, description)
            VALUES
                (:record_type, :amount, :category, :dept_id, :ref_id, :record_date, :description)
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("财务记录新增失败: %s", exc)
            raise

    def get_monthly_sales(self, year: int) -> list[sqlite3.Row]:
        """获取指定年份每月销售收入汇总，用于折线图"""
        sql = """
            SELECT
                strftime('%m', record_date) AS month,
                SUM(amount) AS total
            FROM finance_records
            WHERE record_type='收入'
              AND category='销售收入'
              AND strftime('%Y', record_date) = ?
            GROUP BY month
            ORDER BY month
        """
        return self._conn.execute(sql, (str(year),)).fetchall()

    def get_dept_expense(self, start_date: str, end_date: str) -> list[sqlite3.Row]:
        """获取各部门支出占比，用于饼图"""
        sql = """
            SELECT
                d.dept_name,
                SUM(f.amount) AS total
            FROM finance_records f
            LEFT JOIN departments d ON f.dept_id = d.dept_id
            WHERE f.record_type='支出'
              AND f.record_date BETWEEN ? AND ?
            GROUP BY f.dept_id
            HAVING total > 0
            ORDER BY total DESC
        """
        return self._conn.execute(sql, (start_date, end_date)).fetchall()

    def get_summary(self, start_date: str, end_date: str) -> dict:
        """获取指定时段内总收入、总支出、净利润"""
        sql = """
            SELECT record_type, SUM(amount) AS total
            FROM finance_records
            WHERE record_date BETWEEN ? AND ?
            GROUP BY record_type
        """
        rows = self._conn.execute(sql, (start_date, end_date)).fetchall()
        result = {"收入": 0.0, "支出": 0.0}
        for row in rows:
            result[row["record_type"]] = row["total"] or 0.0
        result["净利润"] = result["收入"] - result["支出"]
        return result

    def get_records(
        self,
        start_date: str = "",
        end_date: str = "",
        record_type: str = "",
        page: int = 1,
        page_size: int = 20,
    ) -> tuple[list[sqlite3.Row], int]:
        """分页查询财务流水"""
        conditions = []
        params: list = []
        if start_date:
            conditions.append("record_date >= ?")
            params.append(start_date)
        if end_date:
            conditions.append("record_date <= ?")
            params.append(end_date)
        if record_type:
            conditions.append("record_type = ?")
            params.append(record_type)

        where = ("WHERE " + " AND ".join(conditions)) if conditions else ""
        total = self._conn.execute(
            f"SELECT COUNT(*) AS cnt FROM finance_records {where}", params
        ).fetchone()["cnt"]

        offset = (page - 1) * page_size
        sql = f"""
            SELECT f.*, d.dept_name
            FROM finance_records f
            LEFT JOIN departments d ON f.dept_id=d.dept_id
            {where}
            ORDER BY record_date DESC, record_id DESC
            LIMIT ? OFFSET ?
        """
        rows = self._conn.execute(sql, params + [page_size, offset]).fetchall()
        return rows, total
