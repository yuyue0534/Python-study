#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TaskFlow —— 全屏任务管理应用（Textual，单文件）

页面（Screen / Mode）
  1 仪表盘   统计卡片、完成趋势、优先级分布、即将到期、最近动态
  2 任务列表 搜索 + 状态/优先级筛选 + 可排序表格 + 详情面板
  3 看板     待办 / 进行中 / 已完成 三列，键盘或鼠标移动卡片
  4 设置     主题、紧凑模式、默认值、数据导出与维护
  ＋ 任务表单页（新建 / 编辑，带校验）、确认对话框、帮助

依赖：  pip install textual           （Python >= 3.9）
运行：  python taskflow.py            # 数据保存在 ~/.taskflow/data.json
        python taskflow.py --data ./my.json
        python taskflow.py --demo     # 内存演示数据，不读写磁盘
按 ? 查看快捷键，Ctrl+P 打开命令面板（可切换主题）。
"""
from __future__ import annotations

import argparse
import csv
import inspect
import json
import os
import random
import sys
import tempfile
import uuid
from collections import Counter
from dataclasses import dataclass, field, fields
from datetime import date, datetime, time as dtime, timedelta
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple

try:
    from rich.text import Text
    from textual import on
    from textual.app import App, ComposeResult
    from textual.binding import Binding
    from textual.containers import Grid, Horizontal, Vertical, VerticalScroll
    from textual.screen import ModalScreen, Screen
    from textual.validation import ValidationResult, Validator
    from textual.widgets import (
        Button, Checkbox, DataTable, Digits, Footer, Header, Input, Label,
        ListItem, ListView, Markdown, OptionList, ProgressBar, RadioButton,
        RadioSet, Select, Sparkline, Static, Switch, TabbedContent, TabPane,
        TextArea,
    )
    from textual.widgets.option_list import Option
except ImportError:  # pragma: no cover
    sys.exit("缺少依赖，请先执行:  pip install textual")

APP_NAME = "TaskFlow"
VERSION = "1.0.0"

# ════════════════════════════════════════════════════════════════════════
# 1. 领域常量与工具函数
# ════════════════════════════════════════════════════════════════════════

STATUSES = [("todo", "待办"), ("doing", "进行中"), ("done", "已完成")]
STATUS_KEYS = [k for k, _ in STATUSES]
STATUS_NAME = dict(STATUSES)
STATUS_COLOR = {"todo": "#f59e0b", "doing": "#3b82f6", "done": "#22c55e"}

PRIORITIES = [("low", "低"), ("medium", "中"), ("high", "高"), ("urgent", "紧急")]
PRIO_NAME = dict(PRIORITIES)
PRIO_RANK = {"urgent": 0, "high": 1, "medium": 2, "low": 3}
PRIO_COLOR = {"low": "#64748b", "medium": "#3b82f6", "high": "#f59e0b", "urgent": "#ef4444"}

WEEKDAYS = "一二三四五六日"


def today() -> date:
    return date.today()


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def parse_date(s: str) -> Optional[date]:
    try:
        return datetime.strptime(s.strip(), "%Y-%m-%d").date()
    except (ValueError, AttributeError):
        return None


def parse_tags(raw: str) -> List[str]:
    seen, out = set(), []
    for part in raw.replace("，", ",").replace("、", ",").split(","):
        p = part.strip().lstrip("#").strip()
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


@dataclass
class Task:
    id: str
    title: str
    desc: str = ""
    status: str = "todo"
    priority: str = "medium"
    due: str = ""                      # YYYY-MM-DD 或空
    tags: List[str] = field(default_factory=list)
    estimate: int = 0                  # 预估工时（小时）
    remind: bool = False
    created: str = field(default_factory=now_iso)
    updated: str = field(default_factory=now_iso)
    done_at: str = ""

    @staticmethod
    def from_dict(d: dict) -> "Task":
        names = {f.name for f in fields(Task)}
        t = Task(**{k: v for k, v in d.items() if k in names})
        if t.status not in STATUS_KEYS:
            t.status = "todo"
        if t.priority not in PRIO_NAME:
            t.priority = "medium"
        t.tags = [str(x) for x in (t.tags or [])]
        t.estimate = int(t.estimate or 0)
        return t

    def due_date(self) -> Optional[date]:
        return parse_date(self.due) if self.due else None

    def days_left(self) -> Optional[int]:
        d = self.due_date()
        return None if d is None else (d - today()).days

    @property
    def overdue(self) -> bool:
        n = self.days_left()
        return self.status != "done" and n is not None and n < 0

    def due_label(self) -> Tuple[str, str]:
        d = self.due_date()
        if d is None:
            return "—", "grey50"
        if self.status == "done":
            return d.strftime("%m-%d"), "grey50"
        n = (d - today()).days
        if n < 0:
            return f"逾期{-n}天", "bold red"
        if n == 0:
            return "今天", "bold yellow"
        if n == 1:
            return "明天", "yellow"
        if n <= 7:
            return f"{n}天后", "cyan"
        return d.strftime("%m-%d"), ""

    def matches(self, q: str) -> bool:
        if not q:
            return True
        hay = " ".join([self.title, self.desc, " ".join(self.tags)]).lower()
        return all(tok in hay for tok in q.lower().split())


@dataclass
class Settings:
    theme: str = "textual-dark"
    username: str = ""
    compact: bool = False
    default_priority: str = "medium"
    confirm_delete: bool = True
    hide_done: bool = False


# ════════════════════════════════════════════════════════════════════════
# 2. 数据层：Store（JSON 持久化，原子写入）
# ════════════════════════════════════════════════════════════════════════

class Store:
    def __init__(self, path: Optional[Path], seed: bool = True):
        self.path = path
        self.tasks: List[Task] = []
        self.activity: List[dict] = []
        self.settings = Settings()
        self.load_error: Optional[str] = None
        self.on_error: Optional[Callable[[str], None]] = None
        loaded = False
        if path is not None and path.exists():
            try:
                self._load()
                loaded = True
            except (OSError, ValueError, KeyError, TypeError, AttributeError) as exc:
                backup = path.with_name(f"{path.name}.corrupt-{datetime.now():%Y%m%d%H%M%S}")
                try:
                    os.replace(path, backup)
                    self.load_error = f"数据文件无法读取（{exc}），已备份为 {backup.name}，将重新开始"
                except OSError:
                    self.load_error = f"数据文件无法读取（{exc}）"
        if not loaded and seed:
            self.seed_demo()
            self.save()

    # ── 持久化 ────────────────────────────────────────────
    def _load(self) -> None:
        raw = json.loads(self.path.read_text(encoding="utf-8"))
        self.tasks = [Task.from_dict(d) for d in raw.get("tasks", [])]
        self.activity = list(raw.get("activity", []))[-200:]
        names = {f.name for f in fields(Settings)}
        self.settings = Settings(**{k: v for k, v in raw.get("settings", {}).items() if k in names})

    def save(self) -> bool:
        if self.path is None:
            return True
        payload = {
            "version": 1,
            "settings": self.settings.__dict__,
            "tasks": [t.__dict__ for t in self.tasks],
            "activity": self.activity[-200:],
        }
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            tmp = self.path.with_name(self.path.name + ".tmp")
            tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            os.replace(tmp, self.path)                 # 原子替换，避免写到一半崩溃损坏数据
            return True
        except OSError as exc:
            if self.on_error:
                self.on_error(f"保存失败：{exc}")
            return False

    # ── 查询 ──────────────────────────────────────────────
    def get(self, task_id: str) -> Optional[Task]:
        return next((t for t in self.tasks if t.id == task_id), None)

    def all_tags(self) -> List[str]:
        return sorted({tag for t in self.tasks for tag in t.tags})

    def stats(self) -> dict:
        tk = self.tasks
        by_status = Counter(t.status for t in tk)
        done_by_day = [0] * 14
        for t in tk:
            if t.done_at:
                try:
                    n = (today() - datetime.fromisoformat(t.done_at).date()).days
                except ValueError:
                    continue
                if 0 <= n < 14:
                    done_by_day[13 - n] += 1
        open_tasks = [t for t in tk if t.status != "done"]
        return {
            "total": len(tk),
            "by_status": by_status,
            "overdue": sum(1 for t in open_tasks if t.overdue),
            "open": len(open_tasks),
            "done_by_day": done_by_day,
            "by_priority": Counter(t.priority for t in open_tasks),
            "hours_open": sum(t.estimate for t in open_tasks),
        }

    def upcoming(self, days: int = 7) -> List[Task]:
        out = [t for t in self.tasks
               if t.status != "done" and t.days_left() is not None and t.days_left() <= days]
        return sorted(out, key=lambda t: (t.due, PRIO_RANK[t.priority]))

    # ── 变更 ──────────────────────────────────────────────
    def log(self, text: str) -> None:
        self.activity.append({"t": now_iso(), "text": text})
        del self.activity[:-200]

    def add(self, **kw) -> Task:
        t = Task(id=uuid.uuid4().hex[:8], **kw)
        if t.status == "done":
            t.done_at = now_iso()
        self.tasks.append(t)
        self.log(f"创建任务「{t.title}」")
        self.save()
        return t

    def update(self, t: Task, **kw) -> None:
        old = t.status
        for k, v in kw.items():
            setattr(t, k, v)
        t.updated = now_iso()
        if t.status == "done" and old != "done":
            t.done_at = now_iso()
            self.log(f"完成任务「{t.title}」")
        elif t.status != "done":
            t.done_at = ""
        if old == t.status:
            self.log(f"更新任务「{t.title}」")
        self.save()

    def set_status(self, task_id: str, status: str) -> None:
        t = self.get(task_id)
        if t and t.status != status:
            self.update(t, status=status)
            if status != "done":
                self.activity[-1]["text"] = f"「{t.title}」→ {STATUS_NAME[status]}"
                self.save()

    def delete(self, task_id: str) -> None:
        t = self.get(task_id)
        if t:
            self.tasks.remove(t)
            self.log(f"删除任务「{t.title}」")
            self.save()

    def clear_done(self) -> int:
        n = len([t for t in self.tasks if t.status == "done"])
        self.tasks = [t for t in self.tasks if t.status != "done"]
        self.log(f"清除了 {n} 个已完成任务")
        self.save()
        return n

    def reset(self) -> None:
        self.tasks.clear()
        self.activity.clear()
        self.log("已清空全部数据")
        self.save()

    # ── 导出 ──────────────────────────────────────────────
    def export(self, kind: str) -> Path:
        out = Path.cwd() / f"taskflow-export-{datetime.now():%Y%m%d-%H%M%S}.{kind}"
        if kind == "json":
            out.write_text(json.dumps([t.__dict__ for t in self.tasks], ensure_ascii=False, indent=2),
                           encoding="utf-8")
        else:
            with open(out, "w", encoding="utf-8-sig", newline="") as f:   # utf-8-sig：Excel 直接打开不乱码
                w = csv.writer(f)
                w.writerow(["ID", "标题", "状态", "优先级", "截止", "标签", "工时", "描述", "创建", "更新", "完成于"])
                for t in self.tasks:
                    w.writerow([t.id, t.title, STATUS_NAME[t.status], PRIO_NAME[t.priority], t.due,
                                ",".join(t.tags), t.estimate, t.desc, t.created, t.updated, t.done_at])
        return out

    # ── 演示数据 ──────────────────────────────────────────
    def seed_demo(self) -> None:
        d = lambda n: (today() + timedelta(days=n)).isoformat()                    # noqa: E731
        ago = lambda n, h=15: datetime.combine(today() - timedelta(days=n), dtime(h, 20)).isoformat(timespec="seconds")  # noqa: E731, E501
        rows = [
            ("设计登录页原型", "完成 Figma 原型并和产品评审。\n\n- 手机号登录\n- 第三方登录入口", "done", "high", -6, ["设计"], 6, 6),
            ("搭建 CI 流水线", "GitHub Actions：lint → test → build → 部署到预发布环境。", "done", "high", -5, ["运维"], 5, 5),
            ("编写 API 接口文档", "使用 OpenAPI 3.0，覆盖用户、订单、支付模块。", "done", "medium", -3, ["文档", "后端"], 8, 3),
            ("修复订单金额舍入错误", "高并发下金额 0.01 偏差，使用 Decimal 重写计算。", "done", "urgent", -2, ["Bug", "后端"], 3, 2),
            ("用户注册流程联调", "前后端联调：校验码、重复注册、错误提示。", "done", "medium", -1, ["前端", "后端"], 4, 1),
            ("数据库慢查询优化", "订单表增加联合索引，目标 P95 < 80ms。\n\n**需要**先在预发布环境验证。", "doing", "urgent", 1, ["后端", "运维"], 6, None),
            ("实现订单列表页", "分页、筛选、导出 CSV。", "doing", "high", 2, ["前端"], 10, None),
            ("接入支付回调验签", "对接两家渠道，注意幂等。", "doing", "high", -1, ["后端"], 8, None),
            ("整理 Q4 路线图", "与产品、运营对齐需求优先级。", "todo", "medium", 5, ["文档"], 3, None),
            ("编写端到端测试用例", "覆盖下单、支付、退款三条主链路。", "todo", "medium", 7, ["测试"], 12, None),
            ("升级 Python 到 3.12", "评估依赖兼容性，先在测试环境试跑。", "todo", "low", 21, ["运维"], 5, None),
            ("设计系统组件梳理", "按钮、表单、表格、弹窗的规范统一。", "todo", "low", 14, ["设计", "前端"], 8, None),
            ("监控告警规则评审", "梳理 P0/P1 告警，减少误报。", "todo", "high", 0, ["运维"], 2, None),
            ("客户反馈问题复盘", "整理上月 12 条反馈，输出改进清单。", "todo", "medium", -2, ["文档"], 2, None),
            ("准备季度技术分享", "主题：Textual 构建终端应用。", "todo", "low", 10, ["文档"], 4, None),
            ("清理无用依赖", "npm 与 pip 依赖各清理一遍。", "todo", "low", 0, ["前端", "后端"], 1, None),
        ]
        for title, desc, st, pr, due, tags, est, done_days_ago in rows:
            t = Task(id=uuid.uuid4().hex[:8], title=title, desc=desc, status=st, priority=pr,
                     due=d(due), tags=tags, estimate=est, remind=(pr == "urgent"))
            t.created = ago(8 + random.randint(0, 3))
            if st == "done":
                t.done_at = ago(done_days_ago or 1)
                t.updated = t.done_at
            self.tasks.append(t)
        self.activity.extend([
            {"t": ago(2), "text": "完成任务「修复订单金额舍入错误」"},
            {"t": ago(1), "text": "完成任务「用户注册流程联调」"},
            {"t": ago(0, 9), "text": "「数据库慢查询优化」→ 进行中"},
        ])


# ════════════════════════════════════════════════════════════════════════
# 3. 表单校验器
# ════════════════════════════════════════════════════════════════════════

class RequiredText(Validator):
    def validate(self, value: str) -> ValidationResult:
        return self.success() if value.strip() else self.failure("标题不能为空")


class DateText(Validator):
    def validate(self, value: str) -> ValidationResult:
        v = value.strip()
        if not v or parse_date(v):
            return self.success()
        return self.failure("日期格式应为 YYYY-MM-DD，例如 2026-12-31")


class HoursText(Validator):
    def validate(self, value: str) -> ValidationResult:
        v = value.strip()
        if not v:
            return self.success()
        if v.isdigit() and 0 <= int(v) <= 999:
            return self.success()
        return self.failure("工时需为 0–999 的整数")


# ════════════════════════════════════════════════════════════════════════
# 4. 通用组件 / 模态框
# ════════════════════════════════════════════════════════════════════════

class ConfirmScreen(ModalScreen[bool]):
    """通用确认框：dismiss(True) 表示确认。"""

    BINDINGS = [
        Binding("y", "yes", "确定", show=False),
        Binding("n,escape", "no", "取消", show=False),
    ]

    def __init__(self, title: str, message: str, yes: str = "确定", danger: bool = False):
        super().__init__()
        self._title, self._message, self._yes, self._danger = title, message, yes, danger

    def compose(self) -> ComposeResult:
        with Vertical(id="dialog", classes="danger" if self._danger else ""):
            yield Label(self._title, id="dialog-title")
            yield Label(self._message, id="dialog-msg")
            with Horizontal(id="dialog-buttons"):
                yield Button(self._yes, variant="error" if self._danger else "primary", id="yes")
                yield Button("取消", id="no")

    def on_mount(self) -> None:
        self.query_one("#no" if self._danger else "#yes", Button).focus()   # 危险操作默认聚焦“取消”

    @on(Button.Pressed)
    def _pressed(self, ev: Button.Pressed) -> None:
        self.dismiss(ev.button.id == "yes")

    def action_yes(self) -> None:
        self.dismiss(True)

    def action_no(self) -> None:
        self.dismiss(False)


HELP_MD = """\
# TaskFlow 快捷键

| 范围 | 按键 | 功能 |
|---|---|---|
| 全局 | `1` `2` `3` `4` | 切换 仪表盘 / 任务 / 看板 / 设置 |
| | `n` | 新建任务 |
| | `Ctrl+P` | 命令面板（切换主题等） |
| | `?` | 本帮助 |
| | `q` / `Ctrl+Q` | 退出 |
| 任务列表 | `/` | 聚焦搜索框 |
| | `↑↓` `Enter` | 选择 / 编辑 |
| | `空格` | 切换完成状态 |
| | `[` `]` | 状态后退 / 前进 |
| | `d` | 删除 |
| | 点击表头 | 按该列排序（再点反向） |
| 看板 | `h` `l` / `←` `→` | 切换列 |
| | `H` `L` / `Shift+←→` | 把卡片移到左 / 右列 |
| | `Enter` `e` / `d` | 编辑 / 删除 |
| 表单 | `Ctrl+S` | 保存 |
| | `Esc` | 取消（有修改会确认） |
| | `Tab` / `Shift+Tab` | 下一个 / 上一个字段 |

按 `Esc` 或 `?` 关闭
"""


class HelpScreen(ModalScreen[None]):
    BINDINGS = [Binding("escape,question_mark,q", "close", "关闭")]

    def compose(self) -> ComposeResult:
        with VerticalScroll(id="help-box"):
            yield Markdown(HELP_MD)

    def action_close(self) -> None:
        self.dismiss(None)


class Sidebar(Vertical):
    PAGES = [("dashboard", "◈  仪表盘"), ("tasks", "☰  任务列表"),
             ("board", "▦  看板"), ("settings", "⚙  设置")]

    def __init__(self, active: str):
        super().__init__(id="sidebar")
        self.active = active

    def compose(self) -> ComposeResult:
        yield Label(f"✦ {APP_NAME}", id="brand")
        yield OptionList(*[Option(label, id=key) for key, label in self.PAGES], id="nav")
        yield Static(id="side-stats")

    def on_mount(self) -> None:
        nav = self.query_one("#nav", OptionList)
        nav.highlighted = [k for k, _ in self.PAGES].index(self.active)
        self.update_stats()

    def update_stats(self) -> None:
        s = self.app.store.stats()
        t = Text()
        t.append(f"未完成 {s['open']}\n", "bold")
        t.append(f"已逾期 {s['overdue']}", "bold red" if s["overdue"] else "grey50")
        self.query_one("#side-stats", Static).update(t)

    @on(OptionList.OptionSelected, "#nav")
    async def _go(self, ev: OptionList.OptionSelected) -> None:
        await self.app.action_switch(ev.option.id)


class PageScreen(Screen):
    """带侧边栏的主页面基类。子类实现 compose_page() 与 refresh_data()。"""

    MODE = ""
    SUB_TITLE = ""

    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        with Horizontal(id="shell"):
            yield Sidebar(self.MODE)
            with Vertical(id="page"):
                yield from self.compose_page()
        yield Footer()

    def compose_page(self) -> ComposeResult:
        raise NotImplementedError

    async def refresh_data(self) -> None:
        pass

    def new_task_status(self) -> Optional[str]:
        return None

    async def sync(self) -> None:
        """数据变化后统一刷新：本页内容 + 侧边栏统计。"""
        r = self.refresh_data()
        if inspect.isawaitable(r):
            await r
        self.query_one(Sidebar).update_stats()

    async def on_screen_resume(self) -> None:
        self.set_class(self.app.store.settings.compact, "compact")
        await self.sync()


# ════════════════════════════════════════════════════════════════════════
# 5. 页面：仪表盘
# ════════════════════════════════════════════════════════════════════════

class DashboardScreen(PageScreen):
    MODE = "dashboard"
    SUB_TITLE = "仪表盘"
    CARDS = [("total", "总任务"), ("doing", "进行中"), ("done", "已完成"), ("overdue", "已逾期")]

    def compose_page(self) -> ComposeResult:
        with VerticalScroll(id="dash"):
            yield Label("", id="greeting", classes="page-title")
            with Grid(id="cards"):
                for key, title in self.CARDS:
                    with Vertical(classes="card", id=f"card-{key}"):
                        yield Label(title, classes="card-title")
                        yield Digits("0", id=f"num-{key}")
                        yield Label("", id=f"sub-{key}", classes="card-sub")
            with Horizontal(id="dash-mid"):
                with Vertical(classes="panel", id="panel-progress"):
                    yield Label("总体完成度", classes="panel-title")
                    yield ProgressBar(total=100, show_eta=False, id="overall")
                    yield Label("近 14 天完成趋势", classes="panel-title")
                    yield Sparkline([0] * 14, id="spark")
                    yield Label("未完成任务 · 优先级分布", classes="panel-title")
                    yield Static(id="prio-bars")
                with Vertical(classes="panel", id="panel-due"):
                    yield Label("即将到期（7 天内）与已逾期    回车编辑", classes="panel-title")
                    yield DataTable(id="due-table", cursor_type="row", zebra_stripes=True)
            with Vertical(classes="panel", id="panel-activity"):
                yield Label("最近动态", classes="panel-title")
                yield Static(id="activity")

    async def refresh_data(self) -> None:
        store = self.app.store
        st = store.stats()
        bs = st["by_status"]
        d = today()
        name = store.settings.username or "朋友"
        hour = datetime.now().hour
        hello = "早上好" if hour < 11 else "中午好" if hour < 14 else "下午好" if hour < 18 else "晚上好"
        self.query_one("#greeting", Label).update(
            f"{hello}，{name}！今天是 {d:%Y年%m月%d日} 星期{WEEKDAYS[d.weekday()]}")

        pct = round(bs["done"] * 100 / st["total"]) if st["total"] else 0
        values = {"total": st["total"], "doing": bs["doing"], "done": bs["done"], "overdue": st["overdue"]}
        subs = {"total": f"待办 {bs['todo']} · 预计剩余 {st['hours_open']} 小时",
                "doing": "正在推进",
                "done": f"完成率 {pct}%",
                "overdue": "需要尽快处理" if st["overdue"] else "一切按时 ✓"}
        for key, _ in self.CARDS:
            self.query_one(f"#num-{key}", Digits).update(str(values[key]))
            self.query_one(f"#sub-{key}", Label).update(subs[key])
        self.query_one("#card-overdue").set_class(st["overdue"] > 0, "alert")

        self.query_one("#overall", ProgressBar).update(total=100, progress=pct)
        self.query_one("#spark", Sparkline).data = list(st["done_by_day"])

        bars = Text()
        peak = max(st["by_priority"].values(), default=1)
        for key, label in reversed(PRIORITIES):
            n = st["by_priority"].get(key, 0)
            bars.append(f"{label:<2} ", "bold")
            bars.append("█" * max(1 if n else 0, round(n * 22 / peak)), PRIO_COLOR[key])
            bars.append(f" {n}\n", "grey70")
        self.query_one("#prio-bars", Static).update(bars)

        table = self.query_one("#due-table", DataTable)
        table.clear(columns=True)
        table.add_column("任务", key="title", width=26)
        table.add_column("截止", key="due", width=9)
        table.add_column("优先级", key="prio", width=6)
        rows = store.upcoming(7)
        for t in rows:
            label, style = t.due_label()
            table.add_row(Text(t.title, no_wrap=True, overflow="ellipsis"), Text(label, style=style),
                          Text(PRIO_NAME[t.priority], style=PRIO_COLOR[t.priority]), key=t.id)
        if not rows:
            table.add_row(Text("🎉 没有临近的截止任务", style="grey50"), "", "", key="none")

        act = Text()
        for item in reversed(store.activity[-7:]):
            try:
                stamp = datetime.fromisoformat(item["t"]).strftime("%m-%d %H:%M")
            except (ValueError, KeyError):
                stamp = "--"
            act.append(f"{stamp}  ", "grey50")
            act.append(item.get("text", "") + "\n")
        self.query_one("#activity", Static).update(act if store.activity else Text("暂无动态", style="grey50"))

    @on(DataTable.RowSelected, "#due-table")
    def _open(self, ev: DataTable.RowSelected) -> None:
        if ev.row_key.value and ev.row_key.value != "none":
            self.app.edit_task(str(ev.row_key.value))


# ════════════════════════════════════════════════════════════════════════
# 6. 页面：任务列表
# ════════════════════════════════════════════════════════════════════════

class TasksScreen(PageScreen):
    MODE = "tasks"
    SUB_TITLE = "任务列表"
    BINDINGS = [
        Binding("slash", "focus_search", "搜索"),
        Binding("escape", "blur_search", "返回表格", show=False),
        Binding("e", "edit", "编辑"),
        Binding("d", "delete", "删除"),
        Binding("space", "toggle_done", "完成/取消"),
        Binding("right_square_bracket", "shift(1)", "状态前进"),
        Binding("left_square_bracket", "shift(-1)", "状态后退", show=False),
    ]
    COLUMNS = [("done", "✓", 3), ("title", "标题", 30), ("status", "状态", 9), ("prio", "优先级", 8),
               ("due", "截止", 9), ("tags", "标签", 14), ("est", "工时", 5)]

    def __init__(self) -> None:
        super().__init__()
        self.query_text = ""
        self.f_status = "all"
        self.f_prio = "all"
        self.sort_key = "due"
        self.sort_rev = False
        self._rebuilding = False

    def compose_page(self) -> ComposeResult:
        with Horizontal(id="filters"):
            yield Input(placeholder="搜索标题 / 描述 / 标签（按 / 聚焦）", id="q")
            yield Select([("全部状态", "all")] + [(n, k) for k, n in STATUSES],
                         value="all", allow_blank=False, id="f-status")
            yield Select([("全部优先级", "all")] + [(n, k) for k, n in PRIORITIES],
                         value="all", allow_blank=False, id="f-prio")
        with Horizontal(id="tasks-body"):
            yield DataTable(id="table", cursor_type="row", zebra_stripes=True)
            with VerticalScroll(id="detail"):
                yield Markdown("", id="detail-md")

    # ── 数据 ──────────────────────────────────────────────
    def _sort_value(self, t: Task):
        k = self.sort_key
        if k == "title":
            return t.title.lower()
        if k == "status":
            return STATUS_KEYS.index(t.status)
        if k == "prio":
            return PRIO_RANK[t.priority]
        if k == "due":
            return (t.due or "9999-99-99", PRIO_RANK[t.priority])
        if k == "tags":
            return ",".join(t.tags)
        if k == "est":
            return t.estimate
        return t.status == "done"

    def visible_tasks(self) -> List[Task]:
        store = self.app.store
        hide = store.settings.hide_done and self.f_status == "all"
        out = [t for t in store.tasks
               if t.matches(self.query_text)
               and (self.f_status == "all" or t.status == self.f_status)
               and (self.f_prio == "all" or t.priority == self.f_prio)
               and not (hide and t.status == "done")]
        return sorted(out, key=self._sort_value, reverse=self.sort_rev)

    def current_id(self) -> Optional[str]:
        table = self.query_one("#table", DataTable)
        if table.row_count == 0:
            return None
        try:
            return str(table.coordinate_to_cell_key(table.cursor_coordinate).row_key.value)
        except Exception:  # noqa: BLE001
            return None

    async def refresh_data(self) -> None:
        table = self.query_one("#table", DataTable)
        keep = self.current_id()
        keep_row = table.cursor_row
        self._rebuilding = True
        try:
            table.clear(columns=True)
            for key, label, width in self.COLUMNS:
                arrow = (" ▼" if self.sort_rev else " ▲") if key == self.sort_key else ""
                table.add_column(Text(label + arrow), key=key, width=width)
            tasks = self.visible_tasks()
            for t in tasks:
                done = t.status == "done"
                label, style = t.due_label()
                table.add_row(
                    Text("✔" if done else "○", style="green" if done else "grey50"),
                    Text(t.title, style="strike grey50" if done else ("bold" if t.priority == "urgent" else ""),
                         no_wrap=True, overflow="ellipsis"),
                    Text("● " + STATUS_NAME[t.status], style=STATUS_COLOR[t.status]),
                    Text(PRIO_NAME[t.priority], style=PRIO_COLOR[t.priority]),
                    Text(label, style=style),
                    Text(" ".join("#" + x for x in t.tags), style="cyan", no_wrap=True, overflow="ellipsis"),
                    Text(f"{t.estimate}h" if t.estimate else "—", style="grey70"),
                    key=t.id)
            if tasks:
                ids = [t.id for t in tasks]
                row = ids.index(keep) if keep in ids else min(keep_row, len(ids) - 1)
                table.move_cursor(row=max(row, 0), animate=False)
        finally:
            self._rebuilding = False
        self.sub_title = f"任务列表（{len(tasks)}/{len(self.app.store.tasks)}）"
        await self._show_detail(self.current_id())

    async def _show_detail(self, task_id: Optional[str]) -> None:
        md = self.query_one("#detail-md", Markdown)
        t = self.app.store.get(task_id) if task_id else None
        if t is None:
            await md.update("*没有匹配的任务*\n\n按 `n` 新建一个吧。")
            return
        label, _ = t.due_label()
        lines = [f"## {t.title}", "",
                 f"**状态** {STATUS_NAME[t.status]}　**优先级** {PRIO_NAME[t.priority]}", "",
                 f"- 截止：{t.due or '未设置'}（{label}）",
                 f"- 预估工时：{t.estimate} 小时" if t.estimate else "- 预估工时：未设置",
                 f"- 标签：{' '.join('`' + x + '`' for x in t.tags) or '无'}",
                 f"- 提醒：{'开启' if t.remind else '关闭'}",
                 f"- 创建：{t.created.replace('T', ' ')}",
                 f"- 更新：{t.updated.replace('T', ' ')}"]
        if t.done_at:
            lines.append(f"- 完成：{t.done_at.replace('T', ' ')}")
        lines += ["", "---", "", t.desc or "*（无描述）*"]
        await md.update("\n".join(lines))

    # ── 事件 ──────────────────────────────────────────────
    @on(Input.Changed, "#q")
    async def _q(self, ev: Input.Changed) -> None:
        self.query_text = ev.value.strip()
        await self.refresh_data()

    @on(Select.Changed, "#f-status")
    async def _fs(self, ev: Select.Changed) -> None:
        self.f_status = str(ev.value)
        await self.refresh_data()

    @on(Select.Changed, "#f-prio")
    async def _fp(self, ev: Select.Changed) -> None:
        self.f_prio = str(ev.value)
        await self.refresh_data()

    @on(DataTable.HeaderSelected, "#table")
    async def _sort(self, ev: DataTable.HeaderSelected) -> None:
        key = str(ev.column_key.value)
        self.sort_rev = (not self.sort_rev) if key == self.sort_key else False
        self.sort_key = key
        await self.refresh_data()

    @on(DataTable.RowHighlighted, "#table")
    async def _hl(self, ev: DataTable.RowHighlighted) -> None:
        if not self._rebuilding and ev.row_key is not None:
            await self._show_detail(str(ev.row_key.value))

    @on(DataTable.RowSelected, "#table")
    def _selected(self, ev: DataTable.RowSelected) -> None:
        self.app.edit_task(str(ev.row_key.value))

    # ── 动作 ──────────────────────────────────────────────
    def action_focus_search(self) -> None:
        self.query_one("#q", Input).focus()

    def action_blur_search(self) -> None:
        self.query_one("#table", DataTable).focus()

    def action_edit(self) -> None:
        if (tid := self.current_id()):
            self.app.edit_task(tid)

    def action_delete(self) -> None:
        if (tid := self.current_id()):
            self.app.delete_task(tid, self.sync)

    async def action_toggle_done(self) -> None:
        if (tid := self.current_id()):
            t = self.app.store.get(tid)
            self.app.store.set_status(tid, "todo" if t.status == "done" else "done")
            await self.sync()

    async def action_shift(self, delta: int) -> None:
        if (tid := self.current_id()):
            t = self.app.store.get(tid)
            i = STATUS_KEYS.index(t.status) + delta
            if 0 <= i < len(STATUS_KEYS):
                self.app.store.set_status(tid, STATUS_KEYS[i])
                await self.sync()
            else:
                self.app.bell()


# ════════════════════════════════════════════════════════════════════════
# 7. 页面：看板
# ════════════════════════════════════════════════════════════════════════

class CardItem(ListItem):
    def __init__(self, task: Task):
        t = Text()
        t.append("● ", PRIO_COLOR[task.priority])
        t.append(task.title, "bold" if task.priority in ("urgent", "high") else "")
        sub = Text()
        label, style = task.due_label()
        if task.due:
            sub.append("⏰ " + label, style)
            sub.append("  ")
        if task.tags:
            sub.append(" ".join("#" + x for x in task.tags), "cyan")
        if task.estimate:
            sub.append(f"  {task.estimate}h", "grey62")
        super().__init__(Static(Text("\n").join([t, sub]) if sub.plain else t))
        self.task_id = task.id


class BoardScreen(PageScreen):
    MODE = "board"
    SUB_TITLE = "看板"
    BINDINGS = [
        Binding("h,left", "col(-1)", "左列", show=False),
        Binding("l,right", "col(1)", "右列", show=False),
        Binding("H,shift+left", "move(-1)", "卡片左移"),
        Binding("L,shift+right", "move(1)", "卡片右移"),
        Binding("e", "edit", "编辑"),
        Binding("d", "delete", "删除"),
    ]

    def compose_page(self) -> ComposeResult:
        with Horizontal(id="board"):
            for key, _ in STATUSES:
                with Vertical(classes="column", id=f"col-{key}"):
                    yield ListView(id=f"list-{key}")

    def _lists(self) -> List[ListView]:
        return [self.query_one(f"#list-{k}", ListView) for k in STATUS_KEYS]

    def _focused_list(self) -> Optional[ListView]:
        f = self.focused
        return f if isinstance(f, ListView) and f in self._lists() else None

    def selected_id(self) -> Optional[str]:
        lv = self._focused_list()
        item = lv.highlighted_child if lv else None
        return getattr(item, "task_id", None)

    def new_task_status(self) -> Optional[str]:
        lv = self._focused_list()
        return STATUS_KEYS[self._lists().index(lv)] if lv else None

    async def refresh_data(self, focus: Optional[str] = None) -> None:
        store = self.app.store
        for key, name in STATUSES:
            lv = self.query_one(f"#list-{key}", ListView)
            prev = getattr(lv.highlighted_child, "task_id", None)
            prev_idx = lv.index or 0
            tasks = sorted((t for t in store.tasks if t.status == key),
                           key=lambda t: (PRIO_RANK[t.priority], t.due or "9999"))
            await lv.clear()
            if tasks:
                await lv.extend([CardItem(t) for t in tasks])
                ids = [t.id for t in tasks]
                want = focus if focus in ids else prev
                lv.index = ids.index(want) if want in ids else min(prev_idx, len(ids) - 1)
            col = self.query_one(f"#col-{key}")
            col.border_title = f"{name}  ({len(tasks)})"
            if focus and focus in [t.id for t in tasks]:
                lv.focus()

    def action_col(self, delta: int) -> None:
        lists = self._lists()
        cur = self._focused_list()
        i = lists.index(cur) if cur else 0
        lists[max(0, min(len(lists) - 1, i + delta))].focus()

    async def action_move(self, delta: int) -> None:
        tid = self.selected_id()
        if not tid:
            return
        t = self.app.store.get(tid)
        i = STATUS_KEYS.index(t.status) + delta
        if not 0 <= i < len(STATUS_KEYS):
            self.app.bell()
            return
        self.app.store.set_status(tid, STATUS_KEYS[i])
        await self.refresh_data(focus=tid)
        self.query_one(Sidebar).update_stats()

    def action_edit(self) -> None:
        if (tid := self.selected_id()):
            self.app.edit_task(tid)

    def action_delete(self) -> None:
        if (tid := self.selected_id()):
            self.app.delete_task(tid, self.sync)

    @on(ListView.Selected)
    def _selected(self, ev: ListView.Selected) -> None:
        if isinstance(ev.item, CardItem):
            self.app.edit_task(ev.item.task_id)


# ════════════════════════════════════════════════════════════════════════
# 8. 页面：设置
# ════════════════════════════════════════════════════════════════════════

class SettingsScreen(PageScreen):
    MODE = "settings"
    SUB_TITLE = "设置"

    def compose_page(self) -> ComposeResult:
        s = self.app.store.settings
        yield Label("设置", classes="page-title")
        with TabbedContent(id="settings-tabs"):
            with TabPane("外观", id="tab-look"):
                with VerticalScroll():
                    yield Label("颜色主题（选择后立即生效，也可用 Ctrl+P 切换）", classes="field-label")
                    yield Select([(n, n) for n in sorted(self.app.available_themes)],
                                 value=self.app.theme, allow_blank=False, id="s-theme")
                    with Horizontal(classes="setting-row"):
                        yield Label("紧凑模式（减少留白，适合小终端）", classes="setting-label")
                        yield Switch(s.compact, id="s-compact")
            with TabPane("行为", id="tab-behavior"):
                with VerticalScroll():
                    yield Label("你的称呼（显示在仪表盘问候语）", classes="field-label")
                    yield Input(s.username, placeholder="例如：小明", id="s-name", max_length=20)
                    yield Label("新建任务的默认优先级", classes="field-label")
                    yield Select([(n, k) for k, n in PRIORITIES], value=s.default_priority,
                                 allow_blank=False, id="s-prio")
                    with Horizontal(classes="setting-row"):
                        yield Label("删除任务前需要确认", classes="setting-label")
                        yield Switch(s.confirm_delete, id="s-confirm")
                    with Horizontal(classes="setting-row"):
                        yield Label("任务列表默认隐藏已完成", classes="setting-label")
                        yield Switch(s.hide_done, id="s-hide")
            with TabPane("数据", id="tab-data"):
                with VerticalScroll():
                    yield Static(id="data-info")
                    with Horizontal(classes="btn-row"):
                        yield Button("导出 CSV", id="b-csv", variant="primary")
                        yield Button("导出 JSON", id="b-json")
                        yield Button("追加演示数据", id="b-demo")
                    with Horizontal(classes="btn-row"):
                        yield Button("清除已完成", id="b-clear", variant="warning")
                        yield Button("重置全部数据", id="b-reset", variant="error")
            with TabPane("关于", id="tab-about"):
                yield Markdown(
                    f"# {APP_NAME} {VERSION}\n\n"
                    "基于 **Textual** 构建的全屏终端任务管理应用。\n\n"
                    "- 多页面：Mode 管理，各页面状态独立保留\n"
                    "- 表单：Input / Select / RadioSet / TextArea / Switch / Checkbox + 校验\n"
                    "- 数据：JSON 持久化，原子写入\n\n按 `?` 查看快捷键。")

    async def refresh_data(self) -> None:
        store = self.app.store
        with self.prevent(Select.Changed):                 # 命令面板切换主题后同步下拉框
            self.query_one("#s-theme", Select).value = self.app.theme
        size = store.path.stat().st_size if store.path and store.path.exists() else 0
        info = Text()
        info.append("数据文件  ", "grey62")
        info.append(f"{store.path or '（内存演示模式，不写磁盘）'}\n")
        info.append("任务数量  ", "grey62")
        info.append(f"{len(store.tasks)}    ")
        info.append("文件大小  ", "grey62")
        info.append(f"{size / 1024:.1f} KB\n")
        self.query_one("#data-info", Static).update(info)

    def _save(self) -> None:
        self.app.store.save()

    @on(Select.Changed, "#s-theme")
    def _theme(self, ev: Select.Changed) -> None:
        if str(ev.value) != self.app.theme:
            self.app.theme = str(ev.value)

    @on(Switch.Changed, "#s-compact")
    def _compact(self, ev: Switch.Changed) -> None:
        self.app.store.settings.compact = ev.value
        self.set_class(ev.value, "compact")
        self._save()

    @on(Input.Changed, "#s-name")
    def _name(self, ev: Input.Changed) -> None:
        self.app.store.settings.username = ev.value.strip()
        self._save()

    @on(Select.Changed, "#s-prio")
    def _prio(self, ev: Select.Changed) -> None:
        self.app.store.settings.default_priority = str(ev.value)
        self._save()

    @on(Switch.Changed, "#s-confirm")
    def _confirm(self, ev: Switch.Changed) -> None:
        self.app.store.settings.confirm_delete = ev.value
        self._save()

    @on(Switch.Changed, "#s-hide")
    def _hide(self, ev: Switch.Changed) -> None:
        self.app.store.settings.hide_done = ev.value
        self._save()

    @on(Button.Pressed)
    async def _button(self, ev: Button.Pressed) -> None:
        store, bid = self.app.store, ev.button.id
        if bid in ("b-csv", "b-json"):
            try:
                path = store.export("csv" if bid == "b-csv" else "json")
                self.notify(f"已导出到 {path}", title="导出成功")
            except OSError as exc:
                self.notify(str(exc), title="导出失败", severity="error")
        elif bid == "b-demo":
            store.seed_demo()
            store.log("追加了演示数据")
            store.save()
            self.notify("已追加 16 条演示任务")
            await self.sync()
        elif bid == "b-clear":
            def done(ok: Optional[bool]) -> None:
                if ok:
                    n = store.clear_done()
                    self.notify(f"已清除 {n} 个已完成任务")
                    self.call_later(self.sync)
            self.app.push_screen(ConfirmScreen("清除已完成任务", "将永久删除所有“已完成”的任务。", "清除", True), done)
        elif bid == "b-reset":
            def done2(ok: Optional[bool]) -> None:
                if ok:
                    store.reset()
                    self.notify("已清空全部数据", severity="warning")
                    self.call_later(self.sync)
            self.app.push_screen(ConfirmScreen("重置全部数据", "所有任务与动态记录都会被删除，且无法恢复！", "全部删除", True), done2)


# ════════════════════════════════════════════════════════════════════════
# 9. 页面：任务表单（独立 Screen，返回 True 表示有保存/删除）
# ════════════════════════════════════════════════════════════════════════

class TaskFormScreen(Screen[bool]):
    BINDINGS = [
        Binding("ctrl+s", "save", "保存"),
        Binding("escape", "cancel", "取消"),
    ]

    def __init__(self, task: Optional[Task] = None, default_status: Optional[str] = None):
        super().__init__()
        self.task = task
        self.default_status = default_status or "todo"
        self._initial: Optional[dict] = None
        self._ready = False

    @property
    def SUB_TITLE(self) -> str:                         # noqa: N802
        return "编辑任务" if self.task else "新建任务"

    def compose(self) -> ComposeResult:
        t, store = self.task, self.app.store
        status = t.status if t else self.default_status
        prio = t.priority if t else store.settings.default_priority
        yield Header(show_clock=True)
        with VerticalScroll(id="form-scroll"):
            yield Label("编辑任务" if t else "新建任务", classes="page-title")
            with Horizontal(id="form-cols"):
                with Vertical(id="form-left"):
                    yield Label("标题 *", classes="field-label")
                    yield Input(t.title if t else "", placeholder="一句话描述要做什么",
                                id="f-title", max_length=120, validators=[RequiredText()])
                    yield Label("", id="e-title", classes="error")
                    yield Label("描述（支持 Markdown）", classes="field-label")
                    yield TextArea(t.desc if t else "", id="f-desc", soft_wrap=True, show_line_numbers=False)
                    yield Label("标签（逗号分隔）", classes="field-label")
                    yield Input(", ".join(t.tags) if t else "", placeholder="例如：后端, Bug", id="f-tags")
                    existing = store.all_tags()
                    if existing:
                        yield Label("已有标签：" + "  ".join("#" + x for x in existing), classes="hint")
                with Vertical(id="form-right"):
                    yield Label("状态", classes="field-label")
                    with RadioSet(id="f-status"):
                        for k, n in STATUSES:
                            yield RadioButton(n, value=(k == status))
                    yield Label("优先级", classes="field-label")
                    yield Select([(n, k) for k, n in PRIORITIES], value=prio, allow_blank=False, id="f-prio")
                    yield Label("截止日期（YYYY-MM-DD，可留空）", classes="field-label")
                    yield Input(t.due if t else "", placeholder="2026-12-31", id="f-due",
                                max_length=10, validators=[DateText()])
                    yield Label("", id="e-due", classes="error")
                    with Horizontal(classes="quick-row"):
                        for bid, text in [("q-today", "今天"), ("q-tomorrow", "明天"),
                                          ("q-week", "下周"), ("q-clear", "清除")]:
                            yield Button(text, id=bid, classes="quick")
                    yield Label("预估工时（小时）", classes="field-label")
                    yield Input(str(t.estimate) if t and t.estimate else "", placeholder="0", id="f-est",
                                type="integer", max_length=3, validators=[HoursText()])
                    yield Label("", id="e-est", classes="error")
                    with Horizontal(classes="setting-row"):
                        yield Label("到期提醒", classes="setting-label")
                        yield Switch(bool(t and t.remind), id="f-remind")
                    if not t:
                        yield Checkbox("保存后继续新建", id="f-again")
            with Horizontal(id="form-buttons"):
                yield Button("保存  Ctrl+S", variant="primary", id="btn-save")
                yield Button("取消  Esc", id="btn-cancel")
                if t:
                    yield Button("删除", variant="error", id="btn-del")
        yield Footer()

    def on_mount(self) -> None:
        self.query_one("#f-title", Input).focus()
        self.call_after_refresh(self._snapshot)

    def _snapshot(self) -> None:
        self._initial = self._collect()
        self._ready = True

    # ── 读取 / 校验 ───────────────────────────────────────
    def _collect(self) -> dict:
        est = self.query_one("#f-est", Input).value.strip()
        idx = self.query_one("#f-status", RadioSet).pressed_index
        return {
            "title": self.query_one("#f-title", Input).value.strip(),
            "desc": self.query_one("#f-desc", TextArea).text.rstrip(),
            "status": STATUS_KEYS[idx if idx >= 0 else 0],
            "priority": str(self.query_one("#f-prio", Select).value),
            "due": self.query_one("#f-due", Input).value.strip(),
            "tags": parse_tags(self.query_one("#f-tags", Input).value),
            "estimate": int(est) if est.isdigit() else 0,
            "remind": self.query_one("#f-remind", Switch).value,
        }

    def _show_error(self, field_id: str, result: Optional[ValidationResult]) -> None:
        err = self.query_one("#e-" + field_id, Label)
        bad = result is not None and not result.is_valid
        err.display = bad
        if bad:
            err.update("✗ " + result.failure_descriptions[0])

    @on(Input.Changed)
    def _changed(self, ev: Input.Changed) -> None:
        if not self._ready or not ev.input.id or ev.input.id not in ("f-title", "f-due", "f-est"):
            return
        self._show_error(ev.input.id[2:], ev.validation_result)

    def _validate_all(self) -> bool:
        ok = True
        first_bad: Optional[Input] = None
        for fid in ("f-title", "f-due", "f-est"):
            inp = self.query_one("#" + fid, Input)
            res = inp.validate(inp.value)
            self._show_error(fid[2:], res)
            if res is not None and not res.is_valid:
                ok = False
                first_bad = first_bad or inp
        if first_bad:
            first_bad.focus()
        return ok

    # ── 动作 ──────────────────────────────────────────────
    def action_save(self) -> None:
        if not self._validate_all():
            self.notify("请先修正标红的字段", severity="warning")
            return
        data, store = self._collect(), self.app.store
        if self.task:
            store.update(self.task, **data)
            self.notify(f"已保存「{data['title']}」")
            self.dismiss(True)
            return
        store.add(**data)
        self.notify(f"已创建「{data['title']}」")
        if self.query_one("#f-again", Checkbox).value:
            self._reset_for_next()
        else:
            self.dismiss(True)

    def _reset_for_next(self) -> None:
        self.query_one("#f-title", Input).value = ""
        self.query_one("#f-desc", TextArea).text = ""
        self.query_one("#f-tags", Input).value = ""
        self.query_one("#f-due", Input).value = ""
        self.query_one("#f-est", Input).value = ""
        self.query_one("#f-remind", Switch).value = False
        for lbl in ("title", "due", "est"):
            self.query_one("#e-" + lbl, Label).display = False
        self.query_one("#f-title", Input).focus()
        self.call_after_refresh(self._snapshot)

    def action_cancel(self) -> None:
        if self._initial is not None and self._collect() != self._initial:
            def after(ok: Optional[bool]) -> None:
                if ok:
                    self.dismiss(False)
            self.app.push_screen(ConfirmScreen("放弃修改？", "有未保存的更改，确定要离开吗？", "放弃", True), after)
        else:
            self.dismiss(False)

    @on(Button.Pressed)
    def _button(self, ev: Button.Pressed) -> None:
        bid = ev.button.id or ""
        if bid == "btn-save":
            self.action_save()
        elif bid == "btn-cancel":
            self.action_cancel()
        elif bid == "btn-del" and self.task:
            self.app.delete_task(self.task.id, lambda: self.dismiss(True))
        elif bid.startswith("q-"):
            days = {"q-today": 0, "q-tomorrow": 1, "q-week": 7}.get(bid)
            self.query_one("#f-due", Input).value = (
                "" if days is None else (today() + timedelta(days=days)).isoformat())


# ════════════════════════════════════════════════════════════════════════
# 10. 应用
# ════════════════════════════════════════════════════════════════════════

class TaskFlowApp(App):
    TITLE = APP_NAME
    MODES = {"dashboard": DashboardScreen, "tasks": TasksScreen,
             "board": BoardScreen, "settings": SettingsScreen}
    DEFAULT_MODE = "dashboard"
    CSS = """
    /* ── 布局骨架 ───────────────────────────────── */
    #shell { height: 1fr; }
    #sidebar { width: 24; background: $panel; padding: 1 0; border-right: tall $primary; }
    #brand { width: 100%; text-align: center; text-style: bold; color: $accent; margin-bottom: 1; }
    #nav { background: transparent; border: none; height: auto; padding: 0; }
    #side-stats { dock: bottom; height: 4; padding: 1 2; color: $text-muted; }
    #page { width: 1fr; padding: 1 2; }
    .page-title { text-style: bold; color: $accent; margin-bottom: 1; }
    .field-label { color: $text-muted; margin-top: 1; }
    .hint { color: $text-muted; margin-top: 1; }
    .error { color: $error; display: none; height: auto; }
    .panel { border: round $primary; padding: 0 1; margin-top: 1; }
    .panel-title { text-style: bold; color: $accent; margin-bottom: 1; }
    .compact #page { padding: 0 1; }
    .compact .field-label { margin-top: 0; }
    .compact .panel { margin-top: 0; }

    /* ── 仪表盘 ─────────────────────────────────── */
    #dash { height: 1fr; }
    #cards { grid-size: 4 1; grid-gutter: 0 2; height: 9; }
    .card { border: round $primary; padding: 0 1; height: 100%; }
    .card-title { color: $text-muted; }
    .card-sub { color: $text-muted; }
    .card Digits { width: auto; height: 3; }
    #card-doing { border: round $accent; }
    #card-done { border: round $success; }
    #card-overdue { border: round $primary; }
    #card-overdue.alert { border: heavy $error; }
    #card-overdue.alert Digits { color: $error; }
    .compact #cards { height: 8; }
    #dash-mid { height: 17; }
    #panel-progress { width: 1fr; }
    #panel-due { width: 1fr; margin-left: 1; }
    #overall { margin-bottom: 1; }
    #spark { height: 3; margin-bottom: 1; }
    #panel-activity { height: 11; }
    #due-table { height: 1fr; }

    /* ── 任务列表 ───────────────────────────────── */
    #filters { height: 3; }
    #filters Input { width: 1fr; }
    #filters Select { width: 20; margin-left: 1; }
    #tasks-body { height: 1fr; margin-top: 1; }
    #table { width: 3fr; }
    #detail { width: 2fr; border-left: tall $primary; padding: 0 1; margin-left: 1; }

    /* ── 看板 ───────────────────────────────────── */
    #board { height: 1fr; }
    .column { width: 1fr; border: round $primary; margin-right: 1; border-title-style: bold; }
    .column:focus-within { border: round $accent; border-title-color: $accent; }
    #col-todo { border-title-color: #f59e0b; }
    #col-doing { border-title-color: #3b82f6; }
    #col-done { border-title-color: #22c55e; }
    .column ListView { background: transparent; height: 1fr; }
    .column ListItem { padding: 0 1; margin-bottom: 1; border-left: tall $panel; }
    .column ListItem.-highlight { border-left: tall $accent; }

    /* ── 设置 ───────────────────────────────────── */
    #settings-tabs { height: 1fr; }
    .setting-row { height: 3; margin-top: 1; align-vertical: middle; }
    .setting-label { width: 1fr; }
    .btn-row { height: auto; margin-top: 1; }
    .btn-row Button { margin-right: 2; }
    #data-info { margin-top: 1; }

    /* ── 表单 ───────────────────────────────────── */
    #form-scroll { padding: 1 4; }
    #form-cols { height: auto; }
    #form-left { width: 3fr; height: auto; margin-right: 3; }
    #form-right { width: 2fr; height: auto; }
    #f-desc { height: 9; }
    #f-status { layout: horizontal; height: auto; border: none; padding: 0; background: transparent; }
    #f-status RadioButton { margin-right: 2; background: transparent; }
    .quick-row { height: 3; }
    Button.quick { min-width: 8; margin-right: 1; }
    #form-buttons { height: auto; margin-top: 2; }
    #form-buttons Button { margin-right: 2; }

    /* ── 模态框 ─────────────────────────────────── */
    ConfirmScreen, HelpScreen { align: center middle; background: $background 60%; }
    #dialog { width: 56; height: auto; padding: 1 2; background: $surface; border: thick $primary; }
    #dialog.danger { border: thick $error; }
    #dialog-title { text-style: bold; margin-bottom: 1; }
    #dialog-msg { margin-bottom: 1; }
    #dialog-buttons { height: auto; align-horizontal: right; }
    #dialog-buttons Button { margin-left: 2; }
    #help-box { width: 76; height: 85%; padding: 1 2; background: $surface; border: thick $primary; }
    """
    BINDINGS = [
        Binding("1", "switch('dashboard')", "仪表盘"),
        Binding("2", "switch('tasks')", "任务"),
        Binding("3", "switch('board')", "看板"),
        Binding("4", "switch('settings')", "设置"),
        Binding("n", "new_task", "新建"),
        Binding("question_mark", "help", "帮助"),
        Binding("q", "quit_app", "退出"),
    ]

    def __init__(self, store: Store):
        super().__init__()
        self.store = store

    def on_mount(self) -> None:
        self.store.on_error = lambda m: self.notify(m, title="错误", severity="error")
        try:
            self.theme = self.store.settings.theme
        except Exception:  # noqa: BLE001  主题不存在时回退默认
            pass
        self.theme_changed_signal.subscribe(self, self._on_theme_changed)
        if self.store.load_error:
            self.notify(self.store.load_error, title="数据恢复", severity="warning", timeout=10)

    def _on_theme_changed(self, theme) -> None:
        if self.store.settings.theme != theme.name:
            self.store.settings.theme = theme.name
            self.store.save()

    # ── 导航 / 全局动作 ───────────────────────────────────
    def _on_page(self) -> bool:
        return isinstance(self.screen, PageScreen)

    async def action_switch(self, mode: str) -> None:
        if self._on_page() and self.current_mode != mode:
            await self.switch_mode(mode)

    def action_new_task(self) -> None:
        if self._on_page():
            self.push_screen(TaskFormScreen(None, self.screen.new_task_status()))

    def action_help(self) -> None:
        if not isinstance(self.screen, (HelpScreen, ConfirmScreen)):
            self.push_screen(HelpScreen())

    def action_quit_app(self) -> None:
        if self._on_page():                       # 表单/对话框打开时不响应，避免误退出丢失输入
            self.exit()

    # ── 业务入口（供各页面复用）─────────────────────────────
    def edit_task(self, task_id: str) -> None:
        t = self.store.get(task_id)
        if t and self._on_page():
            self.push_screen(TaskFormScreen(t))

    def delete_task(self, task_id: str, after: Callable) -> None:
        t = self.store.get(task_id)
        if not t:
            return

        def do_delete() -> None:
            self.store.delete(task_id)
            self.notify(f"已删除「{t.title}」", severity="warning")
            r = after()
            if inspect.isawaitable(r):
                self.call_later(lambda: r)

        if not self.store.settings.confirm_delete:
            do_delete()
            return
        self.push_screen(ConfirmScreen("删除任务", f"确定删除「{t.title}」吗？此操作无法撤销。", "删除", True),
                         lambda ok: do_delete() if ok else None)


# ════════════════════════════════════════════════════════════════════════
# 11. 入口
# ════════════════════════════════════════════════════════════════════════

def default_data_path() -> Path:
    return Path.home() / ".taskflow" / "data.json"


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(prog="taskflow", description=f"{APP_NAME} —— 全屏任务管理（按 ? 查看快捷键）")
    ap.add_argument("--data", type=Path, help=f"数据文件路径（默认 {default_data_path()}）")
    ap.add_argument("--demo", action="store_true", help="使用内存演示数据，不读写磁盘")
    ap.add_argument("--version", action="version", version=f"{APP_NAME} {VERSION}")
    args = ap.parse_args(argv)

    store = Store(None, seed=True) if args.demo else Store(args.data or default_data_path())
    TaskFlowApp(store).run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
