# -*- coding: utf-8 -*-
"""
数据库管理模块 (database/db_manager.py)
职责：
  1. 提供全局唯一的 SQLite 连接（单例模式）
  2. 应用启动时执行建表 DDL，实现数据库自动初始化
  3. 对外暴露 get_connection() 供 DAL 层调用
"""
from __future__ import annotations
import sqlite3
import logging
from config import DB_PATH

logger = logging.getLogger(__name__)


class DatabaseManager:
    """SQLite 数据库管理器（单例）"""

    _instance: "DatabaseManager | None" = None
    _connection: "sqlite3.Connection | None" = None

    def __new__(cls) -> DatabaseManager:
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    # ── 公共接口 ────────────────────────────────────────────────

    def get_connection(self) -> sqlite3.Connection:
        """返回数据库连接，若未初始化则先建立连接"""
        if self._connection is None:
            self._connection = sqlite3.connect(
                DB_PATH,
                check_same_thread=False,   # 允许多线程访问（UI + 后台任务）
                detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES,
            )
            # 返回字典式行对象，方便按列名访问
            self._connection.row_factory = sqlite3.Row
            # 开启外键约束
            self._connection.execute("PRAGMA foreign_keys = ON;")
            logger.info("数据库连接已建立: %s", DB_PATH)
        return self._connection

    def initialize(self) -> None:
        """建立所有数据表（首次运行时调用）"""
        conn = self.get_connection()
        try:
            self._create_tables(conn)
            conn.commit()
            logger.info("数据库初始化完成")
        except sqlite3.Error as exc:
            conn.rollback()
            logger.exception("数据库初始化失败: %s", exc)
            raise

    def close(self) -> None:
        """关闭数据库连接（应用退出时调用）"""
        if self._connection:
            self._connection.close()
            self._connection = None
            logger.info("数据库连接已关闭")

    # ── 私有方法：DDL ───────────────────────────────────────────

    @staticmethod
    def _create_tables(conn: sqlite3.Connection) -> None:
        """执行所有建表语句"""

        ddl_statements = [

            # ── 部门表 ──────────────────────────────────────────
            """
            CREATE TABLE IF NOT EXISTS departments (
                dept_id     TEXT PRIMARY KEY,          -- 部门编号，如 D001
                dept_name   TEXT NOT NULL UNIQUE,      -- 部门名称
                manager     TEXT,                      -- 部门负责人姓名
                parent_id   TEXT,                      -- 上级部门 ID（支持多级架构）
                description TEXT,                      -- 职能描述
                created_at  TEXT DEFAULT (datetime('now','localtime')),
                FOREIGN KEY (parent_id) REFERENCES departments(dept_id)
            );
            """,

            # ── 员工表 ──────────────────────────────────────────
            """
            CREATE TABLE IF NOT EXISTS employees (
                emp_id      TEXT PRIMARY KEY,          -- 工号，如 E001
                name        TEXT NOT NULL,             -- 姓名
                gender      TEXT CHECK(gender IN ('男','女')),
                dept_id     TEXT NOT NULL,             -- 所属部门
                position    TEXT,                      -- 职位
                base_salary REAL DEFAULT 0.0,          -- 基本工资
                hire_date   TEXT,                      -- 入职日期 YYYY-MM-DD
                phone       TEXT,                      -- 联系方式
                status      TEXT DEFAULT '在职' CHECK(status IN ('在职','离职')),
                created_at  TEXT DEFAULT (datetime('now','localtime')),
                FOREIGN KEY (dept_id) REFERENCES departments(dept_id)
            );
            """,

            # ── 商品/物料表 ──────────────────────────────────────
            """
            CREATE TABLE IF NOT EXISTS products (
                product_id   TEXT PRIMARY KEY,         -- 商品编号
                product_name TEXT NOT NULL UNIQUE,
                unit         TEXT DEFAULT '件',        -- 计量单位
                stock_qty    INTEGER DEFAULT 0,        -- 当前库存
                warning_qty  INTEGER DEFAULT 10,       -- 预警阈值
                created_at   TEXT DEFAULT (datetime('now','localtime'))
            );
            """,

            # ── 采购订单表 ──────────────────────────────────────
            """
            CREATE TABLE IF NOT EXISTS purchase_orders (
                po_id        TEXT PRIMARY KEY,
                supplier     TEXT NOT NULL,            -- 供应商名称
                product_id   TEXT NOT NULL,
                quantity     INTEGER NOT NULL CHECK(quantity > 0),
                unit_price   REAL NOT NULL CHECK(unit_price >= 0),
                total_amount REAL GENERATED ALWAYS AS (quantity * unit_price) STORED,
                order_date   TEXT DEFAULT (date('now','localtime')),
                status       TEXT DEFAULT '已入库' CHECK(status IN ('待审核','已入库','已取消')),
                remark       TEXT,
                FOREIGN KEY (product_id) REFERENCES products(product_id)
            );
            """,

            # ── 销售订单表 ──────────────────────────────────────
            """
            CREATE TABLE IF NOT EXISTS sales_orders (
                so_id        TEXT PRIMARY KEY,
                customer     TEXT NOT NULL,            -- 客户名称
                product_id   TEXT NOT NULL,
                quantity     INTEGER NOT NULL CHECK(quantity > 0),
                unit_price   REAL NOT NULL CHECK(unit_price >= 0),
                total_amount REAL GENERATED ALWAYS AS (quantity * unit_price) STORED,
                order_date   TEXT DEFAULT (date('now','localtime')),
                status       TEXT DEFAULT '已完成' CHECK(status IN ('进行中','已完成','已取消')),
                remark       TEXT,
                FOREIGN KEY (product_id) REFERENCES products(product_id)
            );
            """,

            # ── 财务流水表 ──────────────────────────────────────
            """
            CREATE TABLE IF NOT EXISTS finance_records (
                record_id   INTEGER PRIMARY KEY AUTOINCREMENT,
                record_type TEXT NOT NULL CHECK(record_type IN ('收入','支出')),
                amount      REAL NOT NULL CHECK(amount >= 0),
                category    TEXT,                      -- 分类（销售收入/采购支出/员工薪资/其他）
                dept_id     TEXT,                      -- 关联部门
                ref_id      TEXT,                      -- 关联单据号
                record_date TEXT DEFAULT (date('now','localtime')),
                description TEXT,
                FOREIGN KEY (dept_id) REFERENCES departments(dept_id)
            );
            """,

            # ── 初始化预置部门数据 ───────────────────────────────
            """
            INSERT OR IGNORE INTO departments (dept_id, dept_name, manager, description)
            VALUES
                ('D000', '总公司',   '董事长', '集团总部'),
                ('D001', '技术部',   '',       '负责产品研发与技术支持'),
                ('D002', '销售部',   '',       '负责市场开拓与销售'),
                ('D003', '人事部',   '',       '负责招聘与员工管理'),
                ('D004', '财务部',   '',       '负责财务核算与报表');
            """,
        ]

        cursor = conn.cursor()
        for stmt in ddl_statements:
            cursor.execute(stmt)
        logger.debug("所有数据表已创建或已存在")


# ── 模块级单例，供各 DAL 直接 import ───────────────────────────
db_manager = DatabaseManager()
