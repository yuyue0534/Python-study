# -*- coding: utf-8 -*-
"""
员工数据访问层 (dal/employee_dal.py)
仅负责与 employees 表的 CRUD，不含业务校验逻辑。
"""
from __future__ import annotations
import sqlite3
import logging
from typing import Optional
from database import db_manager

logger = logging.getLogger(__name__)


class EmployeeDAL:
    """员工表数据访问对象"""

    def __init__(self):
        self._conn: sqlite3.Connection = db_manager.get_connection()

    # ── 查询 ────────────────────────────────────────────────────

    def get_all(
        self,
        keyword: str = "",
        dept_id: str = "",
        status: str = "在职",
        page: int = 1,
        page_size: int = 20,
    ) -> tuple[list[sqlite3.Row], int]:
        """
        分页查询员工列表，支持姓名/部门模糊查询。
        返回 (行列表, 总记录数)
        """
        conditions = []
        params: list = []

        if keyword:
            conditions.append("(e.name LIKE ? OR d.dept_name LIKE ?)")
            params.extend([f"%{keyword}%", f"%{keyword}%"])
        if dept_id:
            conditions.append("e.dept_id = ?")
            params.append(dept_id)
        if status:
            conditions.append("e.status = ?")
            params.append(status)

        where_clause = ("WHERE " + " AND ".join(conditions)) if conditions else ""

        # 查总数
        count_sql = f"""
            SELECT COUNT(*) AS cnt
            FROM employees e
            LEFT JOIN departments d ON e.dept_id = d.dept_id
            {where_clause}
        """
        total = self._conn.execute(count_sql, params).fetchone()["cnt"]

        # 分页查数据
        offset = (page - 1) * page_size
        data_sql = f"""
            SELECT e.emp_id, e.name, e.gender, e.dept_id, d.dept_name,
                   e.position, e.base_salary, e.hire_date, e.phone, e.status
            FROM employees e
            LEFT JOIN departments d ON e.dept_id = d.dept_id
            {where_clause}
            ORDER BY e.emp_id
            LIMIT ? OFFSET ?
        """
        rows = self._conn.execute(data_sql, params + [page_size, offset]).fetchall()
        return rows, total

    def get_by_id(self, emp_id: str) -> Optional[sqlite3.Row]:
        sql = "SELECT * FROM employees WHERE emp_id = ?"
        return self._conn.execute(sql, (emp_id,)).fetchone()

    def get_max_id(self) -> str:
        """获取当前最大工号，用于自动生成下一个工号"""
        row = self._conn.execute(
            "SELECT emp_id FROM employees ORDER BY emp_id DESC LIMIT 1"
        ).fetchone()
        if row:
            num = int(row["emp_id"][1:]) + 1
        else:
            num = 1
        return f"E{num:04d}"

    # ── 增删改 ──────────────────────────────────────────────────

    def insert(self, data: dict) -> bool:
        """新增员工记录"""
        sql = """
            INSERT INTO employees
                (emp_id, name, gender, dept_id, position, base_salary, hire_date, phone, status)
            VALUES
                (:emp_id, :name, :gender, :dept_id, :position, :base_salary, :hire_date, :phone, :status)
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            logger.info("员工新增成功: %s", data.get("emp_id"))
            return True
        except sqlite3.IntegrityError as exc:
            self._conn.rollback()
            logger.warning("员工新增失败（完整性约束）: %s", exc)
            raise
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("员工新增数据库错误: %s", exc)
            raise

    def update(self, data: dict) -> bool:
        """更新员工信息（以 emp_id 定位）"""
        sql = """
            UPDATE employees SET
                name=:name, gender=:gender, dept_id=:dept_id, position=:position,
                base_salary=:base_salary, hire_date=:hire_date, phone=:phone, status=:status
            WHERE emp_id=:emp_id
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            logger.info("员工更新成功: %s", data.get("emp_id"))
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("员工更新失败: %s", exc)
            raise

    def delete(self, emp_id: str) -> bool:
        """将员工状态改为离职（逻辑删除，保留历史数据）"""
        sql = "UPDATE employees SET status='离职' WHERE emp_id=?"
        try:
            self._conn.execute(sql, (emp_id,))
            self._conn.commit()
            logger.info("员工离职处理: %s", emp_id)
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("员工离职处理失败: %s", exc)
            raise

    def get_by_dept(self, dept_id: str) -> list[sqlite3.Row]:
        """按部门获取所有在职员工"""
        sql = """
            SELECT emp_id, name, position, phone
            FROM employees
            WHERE dept_id=? AND status='在职'
            ORDER BY emp_id
        """
        return self._conn.execute(sql, (dept_id,)).fetchall()
