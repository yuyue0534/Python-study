# -*- coding: utf-8 -*-
"""
部门数据访问层 (dal/department_dal.py)
"""
from __future__ import annotations
import sqlite3
import logging
from typing import Optional
from database import db_manager

logger = logging.getLogger(__name__)


class DepartmentDAL:
    """部门表数据访问对象"""

    def __init__(self):
        self._conn: sqlite3.Connection = db_manager.get_connection()

    def get_all(self) -> list[sqlite3.Row]:
        sql = "SELECT * FROM departments ORDER BY dept_id"
        return self._conn.execute(sql).fetchall()

    def get_by_id(self, dept_id: str) -> Optional[sqlite3.Row]:
        return self._conn.execute(
            "SELECT * FROM departments WHERE dept_id=?", (dept_id,)
        ).fetchone()

    def get_tree(self) -> list[sqlite3.Row]:
        """获取部门层级树（包含父部门信息）"""
        sql = """
            SELECT d.*, p.dept_name AS parent_name
            FROM departments d
            LEFT JOIN departments p ON d.parent_id = p.dept_id
            ORDER BY d.dept_id
        """
        return self._conn.execute(sql).fetchall()

    def get_max_id(self) -> str:
        row = self._conn.execute(
            "SELECT dept_id FROM departments ORDER BY dept_id DESC LIMIT 1"
        ).fetchone()
        if row:
            num = int(row["dept_id"][1:]) + 1
        else:
            num = 1
        return f"D{num:03d}"

    def insert(self, data: dict) -> bool:
        sql = """
            INSERT INTO departments (dept_id, dept_name, manager, parent_id, description)
            VALUES (:dept_id, :dept_name, :manager, :parent_id, :description)
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("部门新增失败: %s", exc)
            raise

    def update(self, data: dict) -> bool:
        sql = """
            UPDATE departments SET
                dept_name=:dept_name, manager=:manager,
                parent_id=:parent_id, description=:description
            WHERE dept_id=:dept_id
        """
        try:
            self._conn.execute(sql, data)
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("部门更新失败: %s", exc)
            raise

    def delete(self, dept_id: str) -> bool:
        """删除部门前需检查是否有员工（BLL 层负责校验）"""
        try:
            self._conn.execute("DELETE FROM departments WHERE dept_id=?", (dept_id,))
            self._conn.commit()
            return True
        except sqlite3.Error as exc:
            self._conn.rollback()
            logger.exception("部门删除失败: %s", exc)
            raise
