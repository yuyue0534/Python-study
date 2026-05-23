# -*- coding: utf-8 -*-
"""
财务与报表可视化模块 (ui/modules/finance_widget.py)
Dashboard：月度销售折线图 + 部门支出饼图 + KPI 卡片 + 流水明细
"""

from datetime import date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTabWidget, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QComboBox, QDateEdit,
    QHeaderView, QFrame, QSizePolicy, QGroupBox
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor

from bll import FinanceBLL
from config import PAGE_SIZE

# Matplotlib 嵌入 PyQt5
import matplotlib
matplotlib.use("Qt5Agg")
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

# 中文字体支持
plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "PingFang SC", "SimHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


# ── KPI 卡片组件 ─────────────────────────────────────────────────

class KpiCard(QFrame):
    """单个 KPI 数据卡片"""

    def __init__(self, title: str, value: str = "¥ 0.00", color: str = "#1565C0"):
        super().__init__()
        self.setObjectName("card_widget")
        self.setMinimumSize(160, 90)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)

        self._lbl_title = QLabel(title)
        self._lbl_title.setStyleSheet("color: #607D8B; font-size: 12px;")

        self._lbl_value = QLabel(value)
        self._lbl_value.setObjectName("label_stat")
        self._lbl_value.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {color};")

        layout.addWidget(self._lbl_title)
        layout.addWidget(self._lbl_value)

    def set_value(self, value: str):
        self._lbl_value.setText(value)


# ── Matplotlib 画布封装 ──────────────────────────────────────────

class ChartCanvas(FigureCanvas):
    """通用 Matplotlib 画布，背景透明融入 Qt 主题"""

    def __init__(self, width=5, height=3.5, dpi=96):
        self.fig = Figure(figsize=(width, height), dpi=dpi, facecolor="#FFFFFF")
        super().__init__(self.fig)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setMinimumHeight(260)


# ════════════════════════════════════════════════════════════════
#  财务报表主界面
# ════════════════════════════════════════════════════════════════

class FinanceWidget(QWidget):
    """财务与报表可视化模块主界面"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._bll          = FinanceBLL()
        self._records_page = 1
        self._build_ui()
        self.refresh_dashboard()

    # ── UI 构建 ─────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)

        tabs = QTabWidget()
        tabs.addTab(self._build_dashboard_tab(), "📊  数据仪表盘")
        tabs.addTab(self._build_records_tab(),   "📋  流水明细")
        layout.addWidget(tabs)

    # ── 仪表盘 Tab ───────────────────────────────────────────────

    def _build_dashboard_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)

        # ── 筛选栏 ─────────────────────────────────────
        filter_bar = QHBoxLayout()
        lbl_year  = QLabel("统计年份:")
        self.combo_year = QComboBox()
        cur_year = date.today().year
        for y in range(cur_year, cur_year - 5, -1):
            self.combo_year.addItem(str(y), y)

        lbl_start = QLabel("支出统计  起:")
        self.date_start = QDateEdit(QDate(cur_year, 1, 1))
        self.date_start.setCalendarPopup(True)
        self.date_start.setDisplayFormat("yyyy-MM-dd")

        lbl_end   = QLabel("止:")
        self.date_end = QDateEdit(QDate.currentDate())
        self.date_end.setCalendarPopup(True)
        self.date_end.setDisplayFormat("yyyy-MM-dd")

        btn_refresh = QPushButton("🔄  刷新图表")
        btn_refresh.clicked.connect(self.refresh_dashboard)

        filter_bar.addWidget(lbl_year)
        filter_bar.addWidget(self.combo_year)
        filter_bar.addSpacing(24)
        filter_bar.addWidget(lbl_start)
        filter_bar.addWidget(self.date_start)
        filter_bar.addWidget(lbl_end)
        filter_bar.addWidget(self.date_end)
        filter_bar.addStretch()
        filter_bar.addWidget(btn_refresh)
        layout.addLayout(filter_bar)

        # ── KPI 卡片行 ─────────────────────────────────
        kpi_row = QHBoxLayout()
        self.kpi_income  = KpiCard("本月收入", color="#2E7D32")
        self.kpi_expense = KpiCard("本月支出", color="#C62828")
        self.kpi_profit  = KpiCard("本月净利润", color="#1565C0")
        self.kpi_warning = KpiCard("库存预警商品", color="#E65100")
        for card in [self.kpi_income, self.kpi_expense, self.kpi_profit, self.kpi_warning]:
            kpi_row.addWidget(card)
        layout.addLayout(kpi_row)

        # ── 图表区 ─────────────────────────────────────
        chart_row = QHBoxLayout()
        chart_row.setSpacing(12)

        # 折线图
        line_group = QGroupBox("月度销售收入趋势（折线图）")
        line_layout = QVBoxLayout(line_group)
        self.line_canvas = ChartCanvas(width=6, height=3.5)
        line_layout.addWidget(self.line_canvas)
        chart_row.addWidget(line_group, 3)

        # 饼图
        pie_group = QGroupBox("各部门支出占比（饼图）")
        pie_layout = QVBoxLayout(pie_group)
        self.pie_canvas = ChartCanvas(width=4, height=3.5)
        pie_layout.addWidget(self.pie_canvas)
        chart_row.addWidget(pie_group, 2)

        layout.addLayout(chart_row)
        return w

    # ── 流水明细 Tab ─────────────────────────────────────────────

    def _build_records_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(12, 12, 12, 12)

        bar = QHBoxLayout()
        self.rec_date_start = QDateEdit(QDate(date.today().year, 1, 1))
        self.rec_date_start.setCalendarPopup(True)
        self.rec_date_start.setDisplayFormat("yyyy-MM-dd")
        self.rec_date_end = QDateEdit(QDate.currentDate())
        self.rec_date_end.setCalendarPopup(True)
        self.rec_date_end.setDisplayFormat("yyyy-MM-dd")
        self.rec_type_combo = QComboBox()
        self.rec_type_combo.addItems(["全部", "收入", "支出"])
        btn_query = QPushButton("查询")
        btn_query.clicked.connect(self._query_records)

        bar.addWidget(QLabel("日期:"))
        bar.addWidget(self.rec_date_start)
        bar.addWidget(QLabel("至"))
        bar.addWidget(self.rec_date_end)
        bar.addWidget(QLabel("类型:"))
        bar.addWidget(self.rec_type_combo)
        bar.addWidget(btn_query)
        bar.addStretch()
        layout.addLayout(bar)

        self.rec_table = QTableWidget()
        self.rec_table.setColumnCount(7)
        self.rec_table.setHorizontalHeaderLabels(
            ["ID", "类型", "金额(元)", "分类", "关联部门", "关联单据", "日期"]
        )
        self.rec_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.rec_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.rec_table.setAlternatingRowColors(True)
        self.rec_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.rec_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.rec_table.verticalHeader().setVisible(False)
        layout.addWidget(self.rec_table)

        page_bar = QHBoxLayout()
        self.rec_btn_prev = QPushButton("◀ 上一页"); self.rec_btn_prev.setObjectName("btn_secondary")
        self.rec_btn_next = QPushButton("下一页 ▶"); self.rec_btn_next.setObjectName("btn_secondary")
        self.rec_lbl_page = QLabel("第 1 页")
        self.rec_btn_prev.clicked.connect(self._rec_prev)
        self.rec_btn_next.clicked.connect(self._rec_next)
        page_bar.addStretch()
        page_bar.addWidget(self.rec_btn_prev)
        page_bar.addWidget(self.rec_lbl_page)
        page_bar.addWidget(self.rec_btn_next)
        layout.addLayout(page_bar)
        return w

    # ── 数据刷新 ─────────────────────────────────────────────────

    def refresh_dashboard(self):
        """刷新全部仪表盘数据：KPI + 折线图 + 饼图"""
        year       = self.combo_year.currentData()
        start_date = self.date_start.date().toString("yyyy-MM-dd")
        end_date   = self.date_end.date().toString("yyyy-MM-dd")

        # KPI
        today    = date.today()
        m_start  = today.replace(day=1).isoformat()
        m_end    = today.isoformat()
        summary  = self._bll.get_summary(m_start, m_end)
        self.kpi_income.set_value(f"¥ {summary['收入']:,.2f}")
        self.kpi_expense.set_value(f"¥ {summary['支出']:,.2f}")
        profit   = summary["净利润"]
        color    = "#2E7D32" if profit >= 0 else "#C62828"
        self.kpi_profit._lbl_value.setStyleSheet(
            f"font-size:22px; font-weight:bold; color:{color};"
        )
        self.kpi_profit.set_value(f"¥ {profit:,.2f}")

        # 库存预警数量
        from bll import InventoryBLL
        inv_bll   = InventoryBLL()
        low_items = inv_bll.get_low_stock()
        self.kpi_warning.set_value(str(len(low_items)) + " 件")

        # 折线图
        self._draw_line_chart(year)

        # 饼图
        self._draw_pie_chart(start_date, end_date)

    def _draw_line_chart(self, year: int):
        """绘制月度销售收入折线图"""
        months, amounts = self._bll.get_monthly_sales_chart_data(year)

        fig = self.line_canvas.fig
        fig.clear()
        ax = fig.add_subplot(111)
        ax.set_facecolor("#FAFCFF")
        fig.patch.set_facecolor("#FFFFFF")

        # 绘制折线 + 面积填充
        ax.plot(months, amounts, color="#1565C0", linewidth=2.5,
                marker="o", markersize=6, markerfacecolor="#FFFFFF",
                markeredgecolor="#1565C0", markeredgewidth=2)
        ax.fill_between(range(len(months)), amounts,
                        alpha=0.12, color="#1565C0")

        # 数据标签
        for i, (m, v) in enumerate(zip(months, amounts)):
            if v > 0:
                ax.annotate(
                    f"¥{v:,.0f}", (i, v),
                    textcoords="offset points", xytext=(0, 8),
                    ha="center", fontsize=8, color="#1565C0"
                )

        ax.set_xticks(range(len(months)))
        ax.set_xticklabels(months, fontsize=9)
        ax.set_title(f"{year} 年月度销售收入（元）", fontsize=11, color="#37474F", pad=10)
        ax.set_ylabel("金额（元）", fontsize=9, color="#607D8B")
        ax.tick_params(colors="#607D8B")
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color("#E0E0E0")
        ax.spines["bottom"].set_color("#E0E0E0")
        ax.yaxis.grid(True, color="#F0F0F0", linestyle="--")
        fig.tight_layout()
        self.line_canvas.draw()

    def _draw_pie_chart(self, start_date: str, end_date: str):
        """绘制各部门支出占比饼图"""
        labels, amounts = self._bll.get_dept_expense_chart_data(start_date, end_date)

        fig = self.pie_canvas.fig
        fig.clear()
        ax = fig.add_subplot(111)
        fig.patch.set_facecolor("#FFFFFF")

        if not amounts:
            ax.text(0.5, 0.5, "暂无支出数据", ha="center", va="center",
                    fontsize=12, color="#9E9E9E", transform=ax.transAxes)
            ax.axis("off")
            self.pie_canvas.draw()
            return

        colors = [
            "#1565C0", "#2E7D32", "#C62828", "#E65100",
            "#6A1B9A", "#00838F", "#F9A825", "#4E342E",
        ]
        wedges, texts, autotexts = ax.pie(
            amounts,
            labels=labels,
            colors=colors[:len(labels)],
            autopct="%1.1f%%",
            startangle=90,
            pctdistance=0.75,
            wedgeprops=dict(width=0.55, edgecolor="white", linewidth=2),
        )
        for t in texts:
            t.set_fontsize(9)
        for at in autotexts:
            at.set_fontsize(8)
            at.set_color("white")
            at.set_fontweight("bold")

        ax.set_title("各部门支出占比", fontsize=11, color="#37474F", pad=10)
        fig.tight_layout()
        self.pie_canvas.draw()

    # ── 流水明细查询 ─────────────────────────────────────────────

    def _query_records(self):
        self._records_page = 1
        self._load_records()

    def _load_records(self):
        start = self.rec_date_start.date().toString("yyyy-MM-dd")
        end   = self.rec_date_end.date().toString("yyyy-MM-dd")
        rtype = self.rec_type_combo.currentText()
        if rtype == "全部":
            rtype = ""

        rows, total, total_pages = self._bll.get_records(
            start_date=start, end_date=end,
            record_type=rtype, page=self._records_page
        )
        self._rec_total_pages = total_pages
        self.rec_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            is_income = row["record_type"] == "收入"
            vals = [
                str(row["record_id"]),
                row["record_type"],
                f"{float(row['amount']):,.2f}",
                row["category"] or "",
                row["dept_name"] or "—",
                row["ref_id"] or "—",
                row["record_date"] or "",
            ]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setTextAlignment(Qt.AlignCenter)
                if c == 1:
                    item.setForeground(
                        QColor("#2E7D32") if is_income else QColor("#C62828")
                    )
                if c == 2:
                    item.setForeground(
                        QColor("#2E7D32") if is_income else QColor("#C62828")
                    )
                self.rec_table.setItem(r, c, item)

        self.rec_lbl_page.setText(
            f"第 {self._records_page} / {total_pages} 页  共 {total} 条"
        )
        self.rec_btn_prev.setEnabled(self._records_page > 1)
        self.rec_btn_next.setEnabled(self._records_page < total_pages)

    def _rec_prev(self):
        if self._records_page > 1:
            self._records_page -= 1; self._load_records()

    def _rec_next(self):
        if self._records_page < getattr(self, "_rec_total_pages", 1):
            self._records_page += 1; self._load_records()
