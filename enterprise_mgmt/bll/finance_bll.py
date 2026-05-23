# -*- coding: utf-8 -*-
"""
财务业务逻辑层 (bll/finance_bll.py)
提供图表所需数据的加工与格式化，供 UI 层直接使用。
"""
from __future__ import annotations
import logging
from datetime import date
from dal import FinanceDAL
from config import PAGE_SIZE

logger = logging.getLogger(__name__)


class FinanceBLL:
    """财务业务逻辑服务"""

    def __init__(self):
        self._dal = FinanceDAL()

    def get_monthly_sales_chart_data(self, year: int) -> tuple[list[str], list[float]]:
        """
        返回用于折线图的 (月份标签列表, 金额列表)
        不存在数据的月份补 0，确保折线图 X 轴完整
        """
        rows = self._dal.get_monthly_sales(year)
        month_map = {row["month"]: float(row["total"]) for row in rows}
        months = [f"{m:02d}月" for m in range(1, 13)]
        amounts = [month_map.get(f"{m:02d}", 0.0) for m in range(1, 13)]
        return months, amounts

    def get_dept_expense_chart_data(
        self, start_date: str, end_date: str
    ) -> tuple[list[str], list[float]]:
        """返回用于饼图的 (部门名称列表, 金额列表)"""
        rows = self._dal.get_dept_expense(start_date, end_date)
        labels  = [r["dept_name"] or "未分配" for r in rows]
        amounts = [float(r["total"]) for r in rows]
        return labels, amounts

    def get_summary(self, start_date: str, end_date: str) -> dict:
        return self._dal.get_summary(start_date, end_date)

    def get_records(
        self,
        start_date: str = "",
        end_date: str = "",
        record_type: str = "",
        page: int = 1,
    ):
        rows, total = self._dal.get_records(
            start_date, end_date, record_type, page, PAGE_SIZE
        )
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        return list(rows), total, total_pages

    @staticmethod
    def current_year() -> int:
        return date.today().year

    @staticmethod
    def default_date_range() -> tuple[str, str]:
        today = date.today()
        start = today.replace(day=1).isoformat()
        return start, today.isoformat()
