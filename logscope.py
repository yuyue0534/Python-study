#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LogScope —— 基于 Textual 的终端日志查看器（单文件）

功能：
  * 自动识别常见日志格式（Python logging / Java logback / syslog / logfmt / JSON Lines）
  * 多行日志（堆栈）自动归并为一条
  * 按级别过滤（1-6 键）、关键字过滤（AND / 排除 / 正则 / 大小写）
  * 实时跟踪（tail -f）、暂停、日志轮转与截断自动处理
  * 虚拟滚动：只渲染可见行，几十万行也流畅
  * 详情面板：完整堆栈、JSON 自动美化
  * 导出过滤结果、复制当前条目、跳转到上/下一个错误

依赖：  pip install textual            （Python >= 3.9）
运行：  python logscope.py /var/log/app.log
        python logscope.py app.log -q "timeout -debug" -l WARNING
        python logscope.py --demo          # 无需日志文件，自带演示数据
按 ? 查看全部快捷键。
"""
from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

try:
    from rich.cells import cell_len
    from rich.console import Group
    from rich.segment import Segment
    from rich.style import Style
    from rich.syntax import Syntax
    from rich.text import Text
    from textual import on
    from textual.app import App, ComposeResult
    from textual.binding import Binding
    from textual.containers import VerticalScroll
    from textual.events import Click
    from textual.geometry import Size
    from textual.message import Message
    from textual.screen import ModalScreen
    from textual.scroll_view import ScrollView
    from textual.strip import Strip
    from textual.widgets import Footer, Header, Input, Static
except ImportError:  # pragma: no cover
    sys.exit("缺少依赖，请先执行:  pip install textual")

VERSION = "1.0.0"

# ════════════════════════════════════════════════════════════════════════
# 1. 常量
# ════════════════════════════════════════════════════════════════════════

LEVELS = ("DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL", "OTHER")
LEVEL_RANK = {lv: i for i, lv in enumerate(LEVELS)}
LEVEL_SHORT = {"DEBUG": "DEBUG", "INFO": "INFO", "WARNING": "WARN",
               "ERROR": "ERROR", "CRITICAL": "CRIT", "OTHER": "OTHER"}
LABEL = {"DEBUG": "DEBUG", "INFO": "INFO ", "WARNING": "WARN ",
         "ERROR": "ERROR", "CRITICAL": "CRIT ", "OTHER": "  ·  "}
LEVEL_COLOR = {"DEBUG": "grey50", "INFO": "green", "WARNING": "yellow",
               "ERROR": "red", "CRITICAL": "bright_magenta", "OTHER": "grey70"}
LEVEL_STYLE = {   # 行内级别标签样式
    "DEBUG": "grey50", "INFO": "bold green", "WARNING": "bold yellow",
    "ERROR": "bold red", "CRITICAL": "bold white on red", "OTHER": "grey50",
}
MSG_STYLE = {     # 行内消息样式
    "DEBUG": "grey58", "INFO": "", "WARNING": "yellow",
    "ERROR": "red", "CRITICAL": "bold red", "OTHER": "grey70",
}
GUTTER_STYLE = "grey37"
TS_STYLE = "cyan"
EXTRA_STYLE = "italic grey50"
HL_STYLE = "bold black on #ffd75f"            # 关键字高亮
CURSOR_BG_FOCUS = "#2f4f7f"
CURSOR_BG_BLUR = "#33363d"

DETAIL_BG = "#0b1220"       # 详情面板背景（与 CSS 保持一致）
GUTTER_WIDTH = 8            # 行号列（7 位 + 空格）
MAX_MSG_CHARS = 2000        # 单行显示的最大字符数（防止超长行拖垮渲染）
MAX_LINES_PER_ENTRY = 2000  # 单条日志最多合并的续行数
READ_CHUNK = 1 << 20        # 每次轮询最多读取 1 MiB

_ALIASES = {
    "TRACE": "DEBUG", "DBG": "DEBUG", "DEBUG": "DEBUG",
    "INFO": "INFO", "NOTICE": "INFO",
    "WARN": "WARNING", "WARNING": "WARNING",
    "ERR": "ERROR", "ERROR": "ERROR", "SEVERE": "ERROR",
    "CRIT": "CRITICAL", "CRITICAL": "CRITICAL", "FATAL": "CRITICAL",
    "EMERG": "CRITICAL", "ALERT": "CRITICAL",
}
_PINO = {10: "DEBUG", 20: "DEBUG", 30: "INFO", 40: "WARNING", 50: "ERROR", 60: "CRITICAL"}
_LV_WORDS = "|".join(sorted(_ALIASES, key=len, reverse=True))

_TS = r"\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:[.,]\d{1,9})?(?:Z|[+-]\d{2}:?\d{2})?"
RE_TS = re.compile(rf"^\[?({_TS})\]?[\s\-|]*")
RE_SYSLOG_TS = re.compile(r"^([A-Z][a-z]{2}\s{1,2}\d{1,2}\s\d{2}:\d{2}:\d{2})\s+")
# 行首级别：INFO / [INFO] / INFO: / INFO - / <INFO>
RE_LEVEL_HEAD = re.compile(rf"^[\[<(]?({_LV_WORDS})\b[\]>)]?:?\s*(?:[-|]\s+)?")
# 级别出现在时间戳后面的任意位置（如 logback 的 "[main] INFO c.e.App"）
RE_LEVEL_ANY = re.compile(rf"(?:^|[\s\[|<(])({_LV_WORDS})(?=[\s\]|:>)])")
RE_LOGFMT_START = re.compile(r"^(?:time|ts|level|t)=")
RE_LOGFMT_KV = re.compile(r'([A-Za-z_][\w.]*)=("(?:[^"\\]|\\.)*"|\S*)')

_TS_KEYS = ("timestamp", "time", "ts", "@timestamp", "asctime", "datetime", "t")
_LV_KEYS = ("level", "levelname", "severity", "lvl", "log_level")
_MSG_KEYS = ("message", "msg", "event", "text")


# ════════════════════════════════════════════════════════════════════════
# 2. 解析层：把文本行变成结构化条目（与 UI 完全解耦，可单独测试）
# ════════════════════════════════════════════════════════════════════════

def norm_level(v) -> str:
    if isinstance(v, bool):
        return "OTHER"
    if isinstance(v, (int, float)):
        return _PINO.get(int(v), "OTHER")
    if isinstance(v, str):
        return _ALIASES.get(v.strip().upper(), "OTHER")
    return "OTHER"


def _clip(s: str, n: int = MAX_MSG_CHARS) -> str:
    return s if len(s) <= n else s[:n] + "…"


def _fmt_ts(v) -> str:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            sec = v / 1000 if v > 1e11 else v
            return datetime.fromtimestamp(sec).strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
        except (OverflowError, OSError, ValueError):
            return str(v)
    return str(v).replace("T", " ", 1)


def _short(v) -> str:
    if isinstance(v, str):
        s = v if re.fullmatch(r"[\w.:/\-@]+", v) else json.dumps(v, ensure_ascii=False)
    else:
        s = json.dumps(v, ensure_ascii=False, separators=(",", ":"), default=str)
    return s if len(s) <= 80 else s[:79] + "…"


def _pick(obj: dict, keys) -> Tuple[Optional[str], object]:
    for k in keys:
        if k in obj:
            return k, obj[k]
    return None, None


def _from_json(obj: dict) -> dict:
    kt, vt = _pick(obj, _TS_KEYS)
    kl, vl = _pick(obj, _LV_KEYS)
    km, vm = _pick(obj, _MSG_KEYS)
    extras = " ".join(f"{k}={_short(v)}" for k, v in obj.items() if k not in (kt, kl, km))
    msg = ("" if km is None else str(vm))
    text = f"{msg}  {extras}".strip() if extras else msg
    return dict(ts=_fmt_ts(vt) if kt else "",
                level=norm_level(vl) if kl else "OTHER",
                msg=_clip(text), data=obj)


def parse_entry_start(line: str) -> Optional[dict]:
    """
    判断一行是否是“新条目的开头”。
    是 → 返回 {ts, level, msg, data}；不是（续行/空行）→ 返回 None。
    """
    s = line.rstrip()
    if not s.strip():
        return None
    st = s.lstrip()

    # JSON Lines
    if st[0] == "{" and st[-1] == "}":
        try:
            obj = json.loads(st)
        except ValueError:
            obj = None
        if isinstance(obj, dict):
            return _from_json(obj)

    # logfmt:  time=... level=info msg="..."
    if RE_LOGFMT_START.match(s):
        kv = {k: v.strip('"') for k, v in RE_LOGFMT_KV.findall(s[:600])}
        lv = kv.get("level") or kv.get("lvl")
        if lv is not None:
            ts = kv.get("time") or kv.get("ts") or kv.get("t") or ""
            msg = re.sub(r'\b(?:time|ts|t|level|lvl)=("[^"]*"|\S+)\s*', "", s, count=3)
            return dict(ts=_fmt_ts(ts) if ts else "", level=norm_level(lv),
                        msg=_clip(msg.strip()), data=None)

    ts, rest = "", s
    m = RE_TS.match(rest) or RE_SYSLOG_TS.match(rest)
    if m:
        ts, rest = m.group(1), rest[m.end():]

    level = None
    m = RE_LEVEL_HEAD.match(rest)
    if m:
        level, rest = _ALIASES[m.group(1)], rest[m.end():]
    elif ts:
        m = RE_LEVEL_ANY.search(rest[:96])
        if m:
            level = _ALIASES[m.group(1)]

    if not ts and level is None:
        return None
    return dict(ts=ts, level=level or "OTHER", msg=_clip(rest.strip()), data=None)


@dataclass
class Entry:
    no: int                      # 首行行号（1 起）
    ts: str
    level: str
    msg: str                     # 首行消息（用于列表展示）
    lines: List[str]             # 全部原始行（首行 + 续行）
    data: Optional[dict] = None  # JSON 条目的原始对象
    blob: str = ""               # 全部行拼接，用于搜索
    extra: int = 0               # 非空续行数
    width: int = 0               # 列表行的显示宽度（列）
    truncated: bool = False

    def __post_init__(self):
        self.blob = "\n".join(self.lines)
        self.width = (GUTTER_WIDTH + (cell_len(self.ts) + 1 if self.ts else 0)
                      + 6 + cell_len(self.msg) + 8)

    def add_line(self, line: str) -> None:
        if len(self.lines) >= MAX_LINES_PER_ENTRY:
            self.truncated = True
            return
        self.lines.append(line)
        self.blob += "\n" + line
        if line.strip():
            self.extra += 1

    @property
    def text(self) -> str:
        return "\n".join(self.lines).rstrip()


class LogStore:
    """条目仓库：逐行喂入，负责多行归并、计数、容量控制。"""

    def __init__(self, max_entries: int = 500_000):
        self.entries: List[Entry] = []
        self.counts: Dict[str, int] = {lv: 0 for lv in LEVELS}
        self.line_no = 0
        self.max_width = 0
        self.max_entries = max_entries

    def clear(self) -> None:
        self.entries.clear()
        self.counts = {lv: 0 for lv in LEVELS}
        self.line_no = 0
        self.max_width = 0

    def feed(self, line: str) -> Optional[Tuple[Entry, bool]]:
        """返回 (条目, 是否新条目)；被忽略的空行返回 None。"""
        self.line_no += 1
        start = parse_entry_start(line)
        if start is None:
            if self.entries:                      # 续行 → 并入上一条
                e = self.entries[-1]
                e.add_line(line)
                return e, False
            if not line.strip():
                return None
            start = dict(ts="", level="OTHER", msg=_clip(line.strip()), data=None)
        e = Entry(no=self.line_no, lines=[line], **start)
        self.entries.append(e)
        self.counts[e.level] += 1
        if e.width > self.max_width:
            self.max_width = e.width
        return e, True

    def trim(self) -> bool:
        """超过容量 10% 时丢弃最旧条目。返回是否发生了裁剪。"""
        if len(self.entries) <= self.max_entries * 1.1:
            return False
        self.entries = self.entries[-self.max_entries:]
        self.counts = {lv: 0 for lv in LEVELS}
        for e in self.entries:
            self.counts[e.level] += 1
        self.max_width = max((e.width for e in self.entries), default=0)
        return True


class Matcher:
    """
    关键字过滤器。
      普通模式：空格分隔的词全部命中（AND）；以 - 开头的词表示排除。
      正则模式：整段输入作为一个正则。
    """

    def __init__(self, text: str = "", regex: bool = False, case: bool = False):
        self.error: Optional[str] = None
        self.inc: List[re.Pattern] = []
        self.exc: List[re.Pattern] = []
        text = text.strip()
        if not text:
            return
        flags = 0 if case else re.IGNORECASE
        try:
            if regex:
                self.inc = [re.compile(text, flags)]
            else:
                for tok in text.split():
                    if tok.startswith("-") and len(tok) > 1:
                        self.exc.append(re.compile(re.escape(tok[1:]), flags))
                    else:
                        self.inc.append(re.compile(re.escape(tok), flags))
        except re.error as e:
            self.error = str(e)
            self.inc, self.exc = [], []

    @property
    def active(self) -> bool:
        return bool(self.inc or self.exc)

    def __call__(self, e: Entry) -> bool:
        if not self.active:
            return True
        b = e.blob
        return all(p.search(b) for p in self.inc) and not any(p.search(b) for p in self.exc)


# ════════════════════════════════════════════════════════════════════════
# 3. IO 层：增量读取（tail -f），处理半行、截断、轮转
# ════════════════════════════════════════════════════════════════════════

class TailReader:
    def __init__(self, path: str, from_end: bool = False):
        self.path = path
        self.fh = None
        self.pos = 0
        self.size = 0
        self.buf = b""
        self._idle_partial = False
        self._open(initial_from_end=from_end)

    @property
    def loading(self) -> bool:
        return self.pos < self.size

    @property
    def progress(self) -> float:
        return 1.0 if self.size == 0 else min(1.0, self.pos / self.size)

    def _open(self, initial_from_end: bool = False) -> None:
        if self.fh:
            self.fh.close()
        self.fh = open(self.path, "rb")
        if initial_from_end:
            self.fh.seek(0, os.SEEK_END)
        self.pos = self.fh.tell()
        self.buf = b""
        self.size = os.fstat(self.fh.fileno()).st_size

    def rewind(self) -> None:
        self.fh.seek(0)
        self.pos, self.buf = 0, b""
        self.size = os.fstat(self.fh.fileno()).st_size

    def close(self) -> None:
        if self.fh:
            self.fh.close()
            self.fh = None

    def read(self, max_bytes: int = READ_CHUNK) -> Tuple[List[str], bool, bool]:
        """返回 (新行列表, 是否还有更多可读, 文件是否被重置)。"""
        reset = False
        st = os.fstat(self.fh.fileno())
        try:
            pst = os.stat(self.path)
        except FileNotFoundError:
            pst = None
        if pst is not None and pst.st_ino and st.st_ino and pst.st_ino != st.st_ino:
            self._open()                       # 日志轮转：路径指向了新文件
            reset, st = True, os.fstat(self.fh.fileno())
        elif st.st_size < self.pos:            # 文件被截断
            self.fh.seek(0)
            self.pos, self.buf = 0, b""
            reset = True
        self.size = st.st_size

        chunk = self.fh.read(max_bytes)
        self.pos = self.fh.tell()
        if not chunk:
            # 连续两次轮询都没有新数据，说明末尾“半行”其实是完整的最后一行
            if self.buf:
                if self._idle_partial:
                    line = self.buf.decode("utf-8", "replace").rstrip("\r")
                    self.buf, self._idle_partial = b"", False
                    return [line], False, reset
                self._idle_partial = True
            return [], False, reset

        self._idle_partial = False
        parts = (self.buf + chunk).split(b"\n")
        self.buf = parts.pop()                 # 最后一段可能是半行，留待下次
        lines = [p.decode("utf-8", "replace").rstrip("\r") for p in parts]
        return lines, len(chunk) >= max_bytes, reset


# ════════════════════════════════════════════════════════════════════════
# 4. UI 组件
# ════════════════════════════════════════════════════════════════════════

def _highlight(text: Text, base: int, plain: str, matcher: Matcher) -> None:
    for p in matcher.inc:
        for m in p.finditer(plain):
            if m.end() > m.start():
                text.stylize(HL_STYLE, base + m.start(), base + m.end())


class LogView(ScrollView, can_focus=True):
    """
    虚拟滚动的日志列表。
    关键点：不为每行创建 Widget，只在 render_line(y) 里为“当前可见的那几十行”现场生成。
    """

    BINDINGS = [
        Binding("up,k", "move(-1)", show=False),
        Binding("down,j", "move(1)", show=False),
        Binding("pageup", "page(-1)", show=False),
        Binding("pagedown", "page(1)", show=False),
        Binding("home,g", "top", show=False),
        Binding("end,G", "bottom", show=False),
    ]

    class Moved(Message):
        """光标移动 / 数据变化后通知上层刷新详情与状态栏。"""

    def __init__(self, **kw):
        super().__init__(**kw)
        self.rows: List[Entry] = []        # 过滤后的可见条目（注意：不能叫 visible，会与 Widget.visible 冲突）
        self.cur = -1
        self.follow = True
        self.matcher = Matcher()
        self._max_width = 0

    # ── 数据 ──────────────────────────────────────────────
    @property
    def selected(self) -> Optional[Entry]:
        return self.rows[self.cur] if 0 <= self.cur < len(self.rows) else None

    def _sync_size(self) -> None:
        self.virtual_size = Size(max(self._max_width, 1), len(self.rows))

    def set_rows(self, rows: List[Entry], max_width: int) -> None:
        """过滤条件变化后整体替换，并尽量让光标停留在原来的条目附近。"""
        keep = self.selected.no if self.selected else None
        self.rows, self._max_width = rows, max_width
        if self.follow or keep is None:
            self.cur = len(rows) - 1
        else:                                  # 二分查找 no >= keep 的第一行
            lo, hi = 0, len(rows)
            while lo < hi:
                mid = (lo + hi) // 2
                if rows[mid].no < keep:
                    lo = mid + 1
                else:
                    hi = mid
            self.cur = min(lo, len(rows) - 1)
        self._sync_size()
        self.call_after_refresh(self._reveal)
        self.refresh()
        self.post_message(self.Moved())

    def append_rows(self, new: List[Entry], max_width: int) -> None:
        self._max_width = max_width
        self.rows.extend(new)
        if self.follow:
            self.cur = len(self.rows) - 1
        elif self.cur < 0 and self.rows:
            self.cur = 0
        self._sync_size()
        self.call_after_refresh(self._reveal)
        self.refresh()
        self.post_message(self.Moved())

    # ── 光标 / 滚动 ──────────────────────────────────────
    def _reveal(self) -> None:
        h = self.size.height
        if h <= 0 or self.cur < 0:
            return
        top = int(self.scroll_offset.y)
        if self.cur < top:
            self.scroll_to(y=self.cur, animate=False)
        elif self.cur >= top + h:
            self.scroll_to(y=self.cur - h + 1, animate=False)

    def goto(self, i: int, follow: Optional[bool] = None) -> None:
        if not self.rows:
            self.cur = -1
            return
        self.cur = max(0, min(i, len(self.rows) - 1))
        if follow is not None:
            self.follow = follow
        self._reveal()
        self.refresh()
        self.post_message(self.Moved())

    def action_move(self, delta: int) -> None:
        self.goto(self.cur + delta, follow=False)

    def action_page(self, direction: int) -> None:
        self.goto(self.cur + direction * max(1, self.size.height - 1), follow=False)

    def action_top(self) -> None:
        self.goto(0, follow=False)

    def action_bottom(self) -> None:
        self.goto(len(self.rows) - 1, follow=True)

    def on_click(self, event: Click) -> None:
        row = event.y + int(self.scroll_offset.y)
        if 0 <= row < len(self.rows):
            self.goto(row, follow=False)

    def on_focus(self) -> None:
        self.refresh()

    def on_blur(self) -> None:
        self.refresh()

    # ── 渲染 ──────────────────────────────────────────────
    def _row_text(self, e: Entry) -> Text:
        t = Text(no_wrap=True, overflow="crop")
        t.append(f"{e.no:>7} ", GUTTER_STYLE)
        if e.ts:
            t.append(e.ts + " ", TS_STYLE)
        t.append(LABEL[e.level] + " ", LEVEL_STYLE[e.level])
        base = len(t.plain)
        t.append(e.msg, MSG_STYLE[e.level] or None)
        _highlight(t, base, e.msg, self.matcher)
        if e.extra:
            t.append(f"  ⏎+{e.extra}", EXTRA_STYLE)
        return t

    def render_line(self, y: int) -> Strip:
        scroll_x, scroll_y = (int(v) for v in self.scroll_offset)
        row = y + scroll_y
        width = self.size.width
        if row >= len(self.rows):
            return Strip.blank(width, self.rich_style)

        e = self.rows[row]
        segments = list(self._row_text(e).render(self.app.console))
        strip = Strip(segments).crop(scroll_x, scroll_x + width)
        if row == self.cur:
            bg = CURSOR_BG_FOCUS if self.has_focus else CURSOR_BG_BLUR
            style = Style(bgcolor=bg)
            return strip.extend_cell_length(width, style).apply_style(style)
        return strip.extend_cell_length(width)


class DetailPane(VerticalScroll):
    """底部详情：完整多行文本、堆栈高亮、JSON 美化。"""

    def compose(self) -> ComposeResult:
        yield Static(id="detail-body")

    def show(self, e: Optional[Entry], matcher: Matcher) -> None:
        body = self.query_one("#detail-body", Static)
        if e is None:
            body.update(Text("（没有可显示的日志）", style="grey50"))
            return
        head = Text()
        head.append(f"#{e.no}  ", "bold")
        if e.ts:
            head.append(e.ts + "  ", TS_STYLE)
        head.append(f" {LEVEL_SHORT[e.level]} ", LEVEL_STYLE[e.level])
        if e.data is not None:
            head.append("  JSON", "italic grey50")
            payload = Syntax(json.dumps(e.data, ensure_ascii=False, indent=2, default=str),
                             "json", theme="ansi_dark", word_wrap=True,
                             background_color=DETAIL_BG)
        elif "Traceback (most recent call last)" in e.blob:
            payload = Syntax(e.text, "pytb", theme="ansi_dark", word_wrap=True,
                             background_color=DETAIL_BG)
        else:
            payload = Text(e.text)
            _highlight(payload, 0, payload.plain, matcher)
        parts = [head, payload]
        if e.truncated:
            parts.append(Text(f"… 已截断，仅保留前 {MAX_LINES_PER_ENTRY} 行", style="italic yellow"))
        body.update(Group(*parts))
        self.scroll_home(animate=False)


class LevelBar(Static):
    def refresh_bar(self, counts: Dict[str, int], on_: Dict[str, bool]) -> None:
        t = Text(" ")
        for i, lv in enumerate(LEVELS):
            label = f" {i + 1}·{LEVEL_SHORT[lv]} {counts[lv]:,} "
            color = LEVEL_COLOR[lv]
            t.append(label, f"bold black on {color}" if on_[lv] else f"{color} dim strike")
            t.append(" ")
        self.update(t)


HELP_TEXT = """\
LogScope 快捷键

 浏览
   ↑ ↓ / j k        上一条 / 下一条
   PgUp PgDn        翻页
   g / Home         跳到开头          G / End   跳到末尾并开启跟随
   ← →              横向滚动（长行）     鼠标     点击选中、滚轮滚动

 过滤
   /                聚焦过滤框（Enter 或 Esc 回到列表）
   1-6              切换 DEBUG / INFO / WARN / ERROR / CRIT / OTHER
   0                重置全部过滤条件
   r                切换 正则 模式
   c                切换 区分大小写
   过滤语法（普通模式）：  timeout redis -debug
                          = 同时含 timeout 和 redis，且不含 debug

 实时
   f                开关 跟随（自动滚到最新）
   p                暂停 / 恢复 读取新日志
   R                从头重新加载文件
   x                清空当前已加载内容

 其它
   n / N            下一个 / 上一个 ERROR 及以上级别
   d                显示 / 隐藏详情面板
   y                复制当前条目
   s                导出当前过滤结果到文件
   ?                本帮助            q   退出

 按 Esc / q / ? 关闭本窗口"""


class HelpScreen(ModalScreen):
    DEFAULT_CSS = """
    HelpScreen { align: center middle; }
    #help-box { width: 76; height: auto; max-height: 95%; padding: 1 2;
                border: round #64748b; background: #0f172a; }
    """
    BINDINGS = [Binding("escape,q,question_mark,f1", "close", "关闭")]

    def compose(self) -> ComposeResult:
        yield Static(Text(HELP_TEXT), id="help-box")

    def action_close(self) -> None:
        self.dismiss(None)


def fmt_size(n: int) -> str:
    f = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if f < 1024 or unit == "GB":
            return f"{f:.0f} {unit}" if unit == "B" else f"{f:.1f} {unit}"
        f /= 1024
    return f"{n} B"


# ════════════════════════════════════════════════════════════════════════
# 5. 应用
# ════════════════════════════════════════════════════════════════════════

class LogScopeApp(App):
    TITLE = "LogScope"
    ENABLE_COMMAND_PALETTE = False
    CSS = """
    #filter { height: 1; border: none; padding: 0 1; background: #1e293b; }
    #filter:focus { background: #334155; }
    #levels { height: 1; background: #0b1220; }
    #log { height: 1fr; scrollbar-size-vertical: 1; scrollbar-size-horizontal: 1; }
    #detail { height: 35%; border-top: solid #475569; padding: 0 1; background: #0b1220; }
    #detail-body { width: 100%; height: auto; }
    #status { height: 1; padding: 0 1; background: #0f172a; }
    """
    BINDINGS = [
        Binding("q", "quit", "退出"),
        Binding("ctrl+q", "quit", "退出", show=False),
        Binding("slash,ctrl+f", "focus_filter", "过滤"),
        Binding("escape", "blur_filter", "返回列表", show=False),
        Binding("f", "toggle_follow", "跟随"),
        Binding("p", "toggle_pause", "暂停"),
        Binding("d", "toggle_detail", "详情"),
        Binding("r", "toggle_regex", "正则"),
        Binding("c", "toggle_case", "大小写"),
        Binding("n", "next_error(1)", "下个错误"),
        Binding("N", "next_error(-1)", "上个错误", show=False),
        Binding("s", "save", "导出"),
        Binding("y", "copy", "复制", show=False),
        Binding("x", "clear", "清空", show=False),
        Binding("R", "reload", "重载", show=False),
        Binding("0", "reset_filters", "重置", show=False),
        Binding("1", "toggle_level(0)", "", show=False),
        Binding("2", "toggle_level(1)", "", show=False),
        Binding("3", "toggle_level(2)", "", show=False),
        Binding("4", "toggle_level(3)", "", show=False),
        Binding("5", "toggle_level(4)", "", show=False),
        Binding("6", "toggle_level(5)", "", show=False),
        Binding("question_mark,f1", "help", "帮助"),
    ]

    def __init__(self, path: str, *, from_end: bool = False, max_entries: int = 500_000,
                 interval: float = 0.5, query: str = "", min_level: Optional[str] = None):
        super().__init__()
        self.path = path
        self.interval = interval
        self.store = LogStore(max_entries)
        self.reader = TailReader(path, from_end=from_end)
        self.level_on = {lv: True for lv in LEVELS}
        if min_level:
            floor = LEVEL_RANK[min_level]
            for lv in LEVELS[:-1]:
                self.level_on[lv] = LEVEL_RANK[lv] >= floor
        self.regex = False
        self.case = False
        self.paused = False
        self.initial_query = query
        self.matcher = Matcher()
        self._debounce = None
        self.sub_title = os.path.abspath(path)

    # ── 布局 ──────────────────────────────────────────────
    def compose(self) -> ComposeResult:
        yield Header(show_clock=True)
        yield Input(value=self.initial_query, placeholder="/ 过滤：空格=AND，-词=排除，r=正则",
                    id="filter")
        yield LevelBar(id="levels")
        yield LogView(id="log")
        yield DetailPane(id="detail")
        yield Static(id="status")
        yield Footer()

    @property
    def view(self) -> LogView:
        return self.query_one("#log", LogView)

    def on_mount(self) -> None:
        self.matcher = Matcher(self.initial_query, self.regex, self.case)
        self.view.matcher = self.matcher
        self.view.focus()
        self.set_interval(self.interval, self._poll)
        self.call_after_refresh(self._poll)

    def on_unmount(self) -> None:
        self.reader.close()

    # ── 数据流 ────────────────────────────────────────────
    def _passes(self, e: Entry) -> bool:
        return self.level_on[e.level] and self.matcher(e)

    def _poll(self) -> None:
        if self.paused:
            self._update_status()
            return
        try:
            lines, more, reset = self.reader.read()
        except OSError as exc:
            self.notify(f"读取失败: {exc}", severity="error")
            return
        if reset:
            self.notify("检测到日志被轮转/截断，已从头读取", severity="warning")
            self.store.clear()
            self._rebuild()
        if lines:
            self._ingest(lines)
        else:
            self._update_status()
        if more:                                  # 大文件分批加载，期间界面保持响应
            self.set_timer(0.005, self._poll)

    def _ingest(self, lines: List[str]) -> None:
        view, store = self.view, self.store
        new_rows: List[Entry] = []
        for line in lines:
            r = store.feed(line)
            if r is None:
                continue
            e, is_new = r
            if is_new:
                if self._passes(e):
                    new_rows.append(e)
            elif self._passes(e):                 # 续行可能让上一条“变得命中”
                last = new_rows[-1] if new_rows else (view.rows[-1] if view.rows else None)
                if last is not e:
                    new_rows.append(e)
        if store.trim():
            self._rebuild()
        else:
            view.append_rows(new_rows, store.max_width)
        self.query_one("#levels", LevelBar).refresh_bar(store.counts, self.level_on)
        self._update_status()

    def _rebuild(self) -> None:
        rows = [e for e in self.store.entries if self._passes(e)]
        self.view.matcher = self.matcher
        self.view.set_rows(rows, self.store.max_width)
        self.query_one("#levels", LevelBar).refresh_bar(self.store.counts, self.level_on)

    def _apply_filter_now(self) -> None:
        self._debounce = None
        text = self.query_one("#filter", Input).value
        self.matcher = Matcher(text, self.regex, self.case)
        self._rebuild()

    # ── 事件 ──────────────────────────────────────────────
    @on(Input.Changed, "#filter")
    def _filter_changed(self) -> None:
        if self._debounce is not None:
            self._debounce.stop()
        self._debounce = self.set_timer(0.2, self._apply_filter_now)

    @on(Input.Submitted, "#filter")
    def _filter_submitted(self) -> None:
        if self._debounce is not None:
            self._debounce.stop()
        self._apply_filter_now()
        self.view.focus()

    @on(LogView.Moved)
    def _moved(self) -> None:
        self.query_one("#detail", DetailPane).show(self.view.selected, self.matcher)
        self._update_status()

    def _update_status(self) -> None:
        v, s = self.view, self.store
        t = Text()
        t.append(f"显示 {len(v.rows):,}/{len(s.entries):,} 条", "bold")
        if v.selected:
            t.append(f"  │  第 {v.selected.no:,} 行")
        t.append("  │  跟随:" + ("开" if v.follow else "关"), "green" if v.follow else "grey50")
        if self.paused:
            t.append("  │  ⏸ 已暂停", "bold yellow")
        if self.regex:
            t.append("  │  正则", "cyan")
        if self.case:
            t.append("  │  区分大小写", "cyan")
        if self.matcher.error:
            t.append(f"  │  正则错误: {self.matcher.error}", "bold red")
        if self.reader.loading and not self.paused:
            t.append(f"  │  加载中 {self.reader.progress:.0%}", "bold yellow")
        t.append(f"  │  {fmt_size(self.reader.size)}", "grey50")
        self.query_one("#status", Static).update(t)

    # ── 动作 ──────────────────────────────────────────────
    def action_focus_filter(self) -> None:
        self.query_one("#filter", Input).focus()

    def action_blur_filter(self) -> None:
        if self.focused is self.query_one("#filter", Input):
            self.view.focus()

    def action_toggle_follow(self) -> None:
        v = self.view
        v.follow = not v.follow
        if v.follow:
            v.goto(len(v.rows) - 1, follow=True)
        self._update_status()

    def action_toggle_pause(self) -> None:
        self.paused = not self.paused
        self._update_status()

    def action_toggle_detail(self) -> None:
        d = self.query_one("#detail", DetailPane)
        d.display = not d.display

    def action_toggle_regex(self) -> None:
        self.regex = not self.regex
        self._apply_filter_now()

    def action_toggle_case(self) -> None:
        self.case = not self.case
        self._apply_filter_now()

    def action_toggle_level(self, idx: int) -> None:
        lv = LEVELS[idx]
        self.level_on[lv] = not self.level_on[lv]
        self._rebuild()

    def action_reset_filters(self) -> None:
        self.level_on = {lv: True for lv in LEVELS}
        inp = self.query_one("#filter", Input)
        inp.value = ""
        self._apply_filter_now()

    def action_next_error(self, direction: int) -> None:
        v = self.view
        i = v.cur + direction
        floor = LEVEL_RANK["ERROR"]
        while 0 <= i < len(v.rows):
            e = v.rows[i]
            if e.level != "OTHER" and LEVEL_RANK[e.level] >= floor:
                v.goto(i, follow=False)
                return
            i += direction
        self.notify("没有更多 ERROR / CRITICAL 了")

    def action_clear(self) -> None:
        self.store.clear()
        self._rebuild()

    def action_reload(self) -> None:
        self.store.clear()
        self.reader.rewind()
        self._rebuild()
        self._poll()

    def action_copy(self) -> None:
        e = self.view.selected
        if not e:
            return
        try:
            self.copy_to_clipboard(e.text)
            self.notify("已复制当前条目（需终端支持 OSC52）")
        except Exception as exc:  # noqa: BLE001
            self.notify(f"复制失败: {exc}", severity="error")

    def action_save(self) -> None:
        rows = self.view.rows
        if not rows:
            self.notify("当前没有可导出的内容", severity="warning")
            return
        stem = os.path.splitext(os.path.basename(self.path))[0] or "log"
        out = f"{stem}.filtered-{datetime.now():%Y%m%d-%H%M%S}.log"
        try:
            with open(out, "w", encoding="utf-8") as f:
                for e in rows:
                    f.write(e.text + "\n")
        except OSError as exc:
            self.notify(f"导出失败: {exc}", severity="error")
            return
        self.notify(f"已导出 {len(rows):,} 条 → {os.path.abspath(out)}")

    def action_help(self) -> None:
        self.push_screen(HelpScreen())


# ════════════════════════════════════════════════════════════════════════
# 6. 演示数据 & CLI
# ════════════════════════════════════════════════════════════════════════

_LOGGERS = ["app.web", "app.db", "app.auth", "app.worker", "app.cache"]
_USERS = ["alice", "bob", "carol", "dave", "测试用户", "erin"]


def _demo_entry(ts: datetime, rnd: random.Random) -> str:
    stamp = f"{ts:%Y-%m-%d %H:%M:%S},{ts.microsecond // 1000:03d}"
    lg = rnd.choice(_LOGGERS)
    ms = rnd.randint(1, 900)
    r = rnd.random()

    def line(level: str, msg: str) -> str:
        return f"{stamp} {level:<8} [{lg}] {msg}"

    if r < 0.34:
        return line("INFO", rnd.choice([
            f"User {rnd.choice(_USERS)} logged in from 10.0.{rnd.randint(0, 9)}.{rnd.randint(1, 250)}",
            f"GET /api/orders/{rnd.randint(1000, 9999)} 200 {ms}ms",
            f"Cache hit ratio {rnd.randint(60, 99)}%",
            f"Job sync-inventory finished in {ms}ms",
        ]))
    if r < 0.55:
        return line("DEBUG", f"SQL: SELECT * FROM orders WHERE id={rnd.randint(1, 5000)} took {ms}ms")
    if r < 0.70:
        return json.dumps({"time": ts.isoformat(timespec="milliseconds"), "level": "info",
                           "msg": "request completed", "path": rnd.choice(["/health", "/api/items", "/login"]),
                           "status": rnd.choice([200, 200, 201, 304]), "latency_ms": ms,
                           "user": {"name": rnd.choice(_USERS), "plan": "pro"}}, ensure_ascii=False)
    if r < 0.82:
        return line("WARNING", rnd.choice([
            f"Slow query detected ({ms + 800}ms)",
            "Retrying connection to redis:6379 (attempt 2/5)",
            f"Disk usage at {rnd.randint(85, 95)}% on /var/lib/data",
        ]))
    if r < 0.95:
        rid = f"{rnd.randint(0, 0xFFFFFF):06x}"
        return (line("ERROR", f"Unhandled exception while processing request id={rid}") + "\n"
                "Traceback (most recent call last):\n"
                '  File "/srv/app/web/handlers.py", line 88, in handle\n'
                "    result = service.process(payload)\n"
                '  File "/srv/app/core/service.py", line 142, in process\n'
                "    conn = self.pool.acquire(timeout=3)\n"
                "redis.exceptions.ConnectionError: Error 111 connecting to redis:6379. Connection refused.")
    if r < 0.98:
        return json.dumps({"time": ts.isoformat(timespec="milliseconds"), "level": "error",
                           "msg": "payment gateway timeout", "order": rnd.randint(1000, 9999),
                           "retry": False, "gateway": {"name": "stripe", "timeout_s": 10}}, ensure_ascii=False)
    return line("CRITICAL", "Out of memory, worker-3 is shutting down")


def make_demo_file(n: int = 4000, live: bool = True) -> str:
    path = os.path.join(tempfile.gettempdir(), "logscope-demo.log")
    rnd = random.Random(42)
    t = datetime.now() - timedelta(minutes=30)
    with open(path, "w", encoding="utf-8") as f:
        for _ in range(n):
            t += timedelta(milliseconds=rnd.randint(50, 900))
            f.write(_demo_entry(t, rnd) + "\n")

    if live:
        def writer() -> None:
            r2 = random.Random()
            while True:
                time.sleep(r2.uniform(0.2, 0.8))
                try:
                    with open(path, "a", encoding="utf-8") as fh:
                        fh.write(_demo_entry(datetime.now(), r2) + "\n")
                except OSError:
                    return
        threading.Thread(target=writer, daemon=True).start()
    return path


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        prog="logscope", description="LogScope —— 终端日志查看器（按 ? 查看快捷键）")
    ap.add_argument("file", nargs="?", help="日志文件路径（省略则进入演示模式）")
    ap.add_argument("--demo", action="store_true", help="使用自动生成并持续追加的演示日志")
    ap.add_argument("-e", "--from-end", action="store_true", help="只显示启动后新增的内容（类似 tail -f）")
    ap.add_argument("-q", "--query", default="", help="初始过滤关键字")
    ap.add_argument("-l", "--level", choices=LEVELS[:-1], type=str.upper, metavar="LEVEL",
                    help="只显示该级别及以上：DEBUG/INFO/WARNING/ERROR/CRITICAL")
    ap.add_argument("--max-entries", type=int, default=500_000, help="内存中最多保留的条目数（默认 500000）")
    ap.add_argument("--interval", type=float, default=0.5, help="轮询新内容的间隔秒数（默认 0.5）")
    ap.add_argument("--version", action="version", version=f"LogScope {VERSION}")
    args = ap.parse_args(argv)

    if args.demo or not args.file:
        path = make_demo_file()
        print(f"[演示模式] 日志文件: {path}", file=sys.stderr)
    else:
        path = args.file
        if not os.path.isfile(path):
            ap.error(f"找不到文件: {path}")
        if not os.access(path, os.R_OK):
            ap.error(f"没有读取权限: {path}")

    LogScopeApp(path, from_end=args.from_end, max_entries=args.max_entries,
                interval=args.interval, query=args.query, min_level=args.level).run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
