# -*- coding: utf-8 -*-
"""
部门业务逻辑层 (bll/department_bll.py)
规则：删除部门前检查是否存在在职员工；部门编号唯一性由 DAL 层数据库约束保障。
"""

import logging
from dal import DepartmentDAL, EmployeeDAL

logger = logging.getLogger(__name__)


class DepartmentBLL:
    """部门业务逻辑服务"""

    def __init__(self):
        self._dal     = DepartmentDAL()
        self._emp_dal = EmployeeDAL()

    def get_all(self):
        return self._dal.get_all()

    def get_tree(self):
        return self._dal.get_tree()

    def get_dept(self, dept_id: str):
        return self._dal.get_by_id(dept_id)

    def add_department(self, form: dict) -> str:
        self._validate(form)
        dept_id = self._dal.get_max_id()
        form["dept_id"] = dept_id
        form.setdefault("parent_id", None)
        self._dal.insert(form)
        logger.info("新增部门: %s %s", dept_id, form.get("dept_name"))
        return dept_id

    def update_department(self, form: dict) -> None:
        if not form.get("dept_id"):
            raise ValueError("部门编号不能为空")
        self._validate(form)
        self._dal.update(form)

    def delete_department(self, dept_id: str) -> None:
        """删除部门前检查是否有在职员工"""
        emps = self._emp_dal.get_by_dept(dept_id)
        if emps:
            names = "、".join(r["name"] for r in emps[:5])
            raise ValueError(
                f"该部门下仍有 {len(emps)} 名在职员工（{names}…），无法删除"
            )
        self._dal.delete(dept_id)
        logger.info("部门已删除: %s", dept_id)

    @staticmethod
    def _validate(form: dict) -> None:
        if not (form.get("dept_name") or "").strip():
            raise ValueError("部门名称不能为空")
        if len(form["dept_name"]) > 50:
            raise ValueError("部门名称不能超过 50 个字符")
