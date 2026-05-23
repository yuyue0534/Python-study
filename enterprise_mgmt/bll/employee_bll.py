# -*- coding: utf-8 -*-
"""
员工业务逻辑层 (bll/employee_bll.py)
负责数据校验、工号生成、联动财务记录（薪资）等业务规则。
DAL 只管 SQL，BLL 管规则。
"""
from __future__ import annotations
import re
import logging
from dal import EmployeeDAL, DepartmentDAL, FinanceDAL
from config import PAGE_SIZE

logger = logging.getLogger(__name__)

# ── 校验常量 ────────────────────────────────────────────────────
PHONE_RE   = re.compile(r"^1[3-9]\d{9}$")
SALARY_MIN = 0
SALARY_MAX = 9_999_999


class EmployeeBLL:
    """员工业务逻辑服务"""

    def __init__(self):
        self._dal      = EmployeeDAL()
        self._dept_dal = DepartmentDAL()
        self._fin_dal  = FinanceDAL()

    # ── 查询 ────────────────────────────────────────────────────

    def get_employee_list(
        self,
        keyword: str = "",
        dept_id: str = "",
        status: str = "在职",
        page: int = 1,
    ) -> tuple[list, int, int]:
        """
        返回 (行列表, 总记录数, 总页数)
        """
        rows, total = self._dal.get_all(
            keyword=keyword,
            dept_id=dept_id,
            status=status,
            page=page,
            page_size=PAGE_SIZE,
        )
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        return list(rows), total, total_pages

    def get_employee(self, emp_id: str):
        return self._dal.get_by_id(emp_id)

    def get_employees_by_dept(self, dept_id: str):
        return self._dal.get_by_dept(dept_id)

    # ── 新增 ────────────────────────────────────────────────────

    def add_employee(self, form: dict) -> str:
        """
        新增员工。
        返回新分配的工号；抛出 ValueError 表示校验失败。
        """
        self._validate(form)

        # 自动生成工号
        emp_id = self._dal.get_max_id()
        form["emp_id"] = emp_id
        form.setdefault("status", "在职")

        self._dal.insert(form)
        logger.info("员工 %s (%s) 入职登记完成", form["name"], emp_id)
        return emp_id

    # ── 修改 ────────────────────────────────────────────────────

    def update_employee(self, form: dict) -> None:
        """修改员工信息"""
        if not form.get("emp_id"):
            raise ValueError("工号不能为空")
        self._validate(form)
        self._dal.update(form)

    # ── 离职 ────────────────────────────────────────────────────

    def resign_employee(self, emp_id: str) -> None:
        """办理员工离职（逻辑删除）"""
        emp = self._dal.get_by_id(emp_id)
        if not emp:
            raise ValueError(f"工号 {emp_id} 不存在")
        if emp["status"] == "离职":
            raise ValueError(f"员工 {emp['name']} 已办理离职")
        self._dal.delete(emp_id)
        logger.info("员工 %s 已办理离职", emp_id)

    # ── 内部校验 ────────────────────────────────────────────────

    def _validate(self, form: dict) -> None:
        """统一校验入参，校验失败抛出 ValueError"""
        name = (form.get("name") or "").strip()
        if not name:
            raise ValueError("姓名不能为空")
        if len(name) > 50:
            raise ValueError("姓名不能超过 50 个字符")

        dept_id = form.get("dept_id")
        if not dept_id:
            raise ValueError("请选择所属部门")
        if not self._dept_dal.get_by_id(dept_id):
            raise ValueError(f"部门 {dept_id} 不存在")

        salary = form.get("base_salary")
        try:
            salary = float(salary)
        except (TypeError, ValueError):
            raise ValueError("基本工资必须为数字")
        if not (SALARY_MIN <= salary <= SALARY_MAX):
            raise ValueError(f"基本工资须在 {SALARY_MIN} ~ {SALARY_MAX} 之间")
        form["base_salary"] = salary

        phone = (form.get("phone") or "").strip()
        if phone and not PHONE_RE.match(phone):
            raise ValueError("手机号格式不正确（11位大陆手机号）")

        hire_date = (form.get("hire_date") or "").strip()
        if not hire_date:
            raise ValueError("入职日期不能为空")
