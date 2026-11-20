"""Timesheet Merger fixed（Python 3.10+；依赖 openpyxl）

自动读取脚本旁 teaminfoList.txt，不再需要剪贴板或回车。
安装：python -m pip install openpyxl
运行：python format_timesheet_merger-fixed.py
指定：python format_timesheet_merger-fixed.py --staff teaminfoList.txt --src ./2026Nov2 --start 2026-11-09 --end 2026-11-15
TXT应为Excel复制的Tab分隔文本；原A~L列或SID/Function/Workstream/Subteam/Sub-Subteam五列。
日期与默认目录沿用上传版本，可在配置区修改；相对路径基于脚本目录。
"""

import argparse
import logging
import math
import re
import sys
import os
import tempfile
from collections import Counter
from decimal import Decimal
from datetime import datetime, date, timedelta
from pathlib import Path

import csv
import io

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side





# ─────────────────────────────────────────────
# 配置区（按需修改）
# ─────────────────────────────────────────────
DEFAULT_SRC_DIR = "./2026Nov2"
DEFAULT_OUT_DIR = "./output"

# 输出固定列定义（顺序即为输出顺序）
# 示例人员工号：对应 sheet 整页跳过，不提取工时
EXCLUDED_STAFF_IDS = {"XCS01818"}
# 仅保留 Date 在此范围内的明细，包含开始和结束日期
FILTER_START_DATE = "2026-11-09"
FILTER_END_DATE = "2026-11-15"

# 相对路径默认以脚本所在目录为基准，避免从其他目录启动时找不到文件。
BASE_DIR = Path(__file__).resolve().parent
DEFAULT_STAFF_FILE = "teaminfoList.txt"
STAFF_ENCODING = "auto"  # auto / utf-8-sig / utf-16 / gb18030 等
TEXT_DATE_ORDER = "DMY"  # 原脚本优先日/月/年；如为美式，改为 MDY
PERSON_SCAN_ROWS = 10
HEADER_SCAN_ROWS = 15
ALLOW_NUMERIC_ID_ZERO_ALIAS = True  # 仅纯数字工号允许去前导零兜底；冲突别名禁用
AUDIT_EVENTS = []

# Staff 花名册列位置（0-indexed，A=0）
STAFF_COL_POSITIONS = [0, 8, 9, 10, 11]          # A, I, J, K, L
STAFF_COL_NAMES     = ["SID", "Function", "Workstream", "Subteam", "Sub-Subteam"]
STAFF_ID_COL        = "SID"

# 输出固定列定义（顺序即为输出顺序）
OUTPUT_COLS = [
    "Date", "Role/Type", "Task Nature", "Module", "Hours",
    "Ref ID", "Staff ID", "Name", "Function", "Workstream", "Subteam", "Sub-Subteam"
]

# 全局 logger 占位，由 main() 调用 setup_logger() 后正式初始化
logger: logging.Logger = logging.getLogger("timesheet")

# ─────────────────────────────────────────────
# 日志初始化（同时输出控制台 + log 文件）
# ─────────────────────────────────────────────
def setup_logger(log_path: Path) -> logging.Logger:
    log = logging.getLogger("timesheet")
    for handler in list(log.handlers):
        log.removeHandler(handler)
        handler.close()
    log.setLevel(logging.INFO)
    log.propagate = False
    fmt = logging.Formatter("%(asctime)s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    for handler in (logging.StreamHandler(), logging.FileHandler(log_path, encoding="utf-8")):
        handler.setFormatter(fmt)
        log.addHandler(handler)
    return log


def clean_text(value) -> str:
    return "" if value is None else str(value).strip()


def normalize_staff_id(value) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        if not math.isfinite(value) or value != int(value):
            return ""
        return str(int(value))
    return str(value).strip().upper()


def label_text(value) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip().rstrip(":：").strip().casefold()


def audit(path, sheet, row, reason, detail=""):
    AUDIT_EVENTS.append({"File": str(path), "Sheet": sheet, "Row": row,
                         "Reason": reason, "Detail": str(detail)})
    logger.warning(f"  [{reason}] {Path(path).name} / {sheet} / 行 {row}：{detail}")


class StaffRoster(dict):
    """精确工号优先；只在没有歧义时提供纯数字去零别名。"""
    def __init__(self):
        super().__init__()
        self.aliases = {}
        self.alias_owners = {}
        self.ambiguous = set()

    def build_aliases(self):
        if not ALLOW_NUMERIC_ID_ZERO_ALIAS:
            return
        groups = {}
        for sid in self:
            if sid.isdecimal():
                groups.setdefault(sid.lstrip("0") or "0", []).append(sid)
        for alias, owners in groups.items():
            if len(owners) == 1:
                self.aliases[alias] = self[owners[0]]
                self.alias_owners[alias] = owners[0]
            else:
                self.ambiguous.add(alias)
                logger.warning(f"[花名册] 禁用冲突的数字别名 {alias}：{', '.join(owners)}；精确工号仍可匹配")

    def get(self, key, default=None):
        key = normalize_staff_id(key)
        if key in self:
            return super().get(key)
        if ALLOW_NUMERIC_ID_ZERO_ALIAS and key.isdecimal():
            return self.aliases.get(key.lstrip("0") or "0", default)
        return default


def load_staff_from_txt(path: Path, encoding: str = STAFF_ENCODING) -> StaffRoster:
    """读取 Excel 复制保存的 Tab 文本；支持原 A~L 或精简的五列格式。"""
    data = path.read_bytes()
    if encoding != "auto":
        text = data.decode(encoding)
        used_encoding = encoding
    else:
        if data.startswith((b"\xff\xfe", b"\xfe\xff")):
            encodings = ["utf-16"]
        elif b"\x00" in data[:200]:
            raise ValueError("TXT 可能是无 BOM 的 UTF-16，请通过 --staff-encoding utf-16-le 或 utf-16-be 指定编码")
        else:
            encodings = ["utf-8-sig", "gb18030"]
        for used_encoding in encodings:
            try:
                text = data.decode(used_encoding)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("无法识别 TXT 编码，请用 --staff-encoding 指定，或另存为 UTF-8")
    text = text.lstrip("\ufeff")
    if "\x00" in text:
        raise ValueError("TXT 含空字符，请检查编码")
    reader = csv.reader(io.StringIO(text, newline=""), delimiter="\t", strict=True)
    rows = []
    for row in reader:
        if any(clean_text(cell) for cell in row):
            rows.append((reader.line_num, row))
    if not rows:
        raise ValueError("花名册 TXT 为空")
    first = [label_text(v) for v in rows[0][1]]
    required = [label_text(v) for v in STAFF_COL_NAMES]
    if all(first.count(v) == 1 for v in required):
        positions = [first.index(v) for v in required]
        rows = rows[1:]
    elif first and first[0] in {"sid", "staff id", "工号"}:
        if len(first) < 12:
            raise ValueError("花名册标题无法映射五个字段，也不足 A~L 共12列")
        positions = STAFF_COL_POSITIONS
        rows = rows[1:]
        logger.warning("[花名册] 标题未完整匹配，按原 A/I/J/K/L 列位置读取")
    else:
        width = len(rows[0][1])
        if width == 5:
            positions = list(range(5))
        elif width >= 12:
            positions = STAFF_COL_POSITIONS
        else:
            raise ValueError(f"花名册应为 Tab 分隔的12列或5列，首条记录只有 {width} 列；不能用普通空格代替 Tab")
        logger.warning("[花名册] 未识别到标题，按列位置读取，首行作为员工数据")
    roster = StaffRoster()
    for line_num, row in rows:
        if len(row) <= max(positions):
            raise ValueError(f"花名册第 {line_num} 行列数不足：{len(row)} 列，至少需要 {max(positions)+1} 列；请保留空列的 Tab")
        values = [clean_text(row[i]) for i in positions]
        if label_text(values[0]) in {"sid", "staff id", "工号"}:
            continue
        sid = normalize_staff_id(values[0])
        if not sid:
            raise ValueError(f"花名册第 {line_num} 行缺少工号")
        info = dict(zip(STAFF_COL_NAMES[1:], [v or "-" for v in values[1:]]))
        if sid in roster:
            if roster[sid] != info:
                raise ValueError(f"花名册第 {line_num} 行工号 {sid} 重复且部门信息冲突")
            logger.warning(f"[花名册] 第 {line_num} 行工号 {sid} 完全重复，保留一条")
            continue
        roster[sid] = info
    if not roster:
        raise ValueError("花名册没有有效员工")
    roster.build_aliases()
    logger.info(f"[花名册] 自动读取 {path}；编码 {used_encoding}；{len(roster)} 位员工")
    return roster


def discover_files(src_dir: Path) -> list[Path]:
    files = sorted((p for p in src_dir.iterdir()
                    if p.is_file() and p.suffix.lower() == ".xlsx"
                    and not p.name.startswith("~$")
                    and not p.name.lower().startswith("output_")), key=lambda p: p.name.casefold())
    if not files:
        raise FileNotFoundError(f"在 {src_dir} 未找到源 .xlsx 文件（不递归，忽略临时文件和 output_ 输出）")
    logger.info(f"[发现] 共 {len(files)} 个源文件")
    return files


# ─────────────────────────────────────────────
# Step 2：日期处理工具
# ─────────────────────────────────────────────
def _parse_date(val, epoch=None) -> str | None:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        if not math.isfinite(val) or val < 1:
            return None
        try:
            from openpyxl.utils.datetime import from_excel, WINDOWS_EPOCH
            result = from_excel(val, epoch=epoch or WINDOWS_EPOCH)
            return result.strftime("%Y-%m-%d") if isinstance(result, (datetime, date)) else None
        except (ValueError, OverflowError, TypeError):
            return None
    value = clean_text(val)
    # 英文月份显式映射，避免 Windows 区域设置导致 %b 解析失败
    match = re.fullmatch(r"(\d{1,2})-([A-Za-z]{3})-(\d{2}|\d{4})", value)
    if match:
        months = {m: i for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split(), 1)}
        day, month, year = match.groups()
        year = int(year)
        if len(match.group(3)) == 2:
            year += 2000 if year <= 68 else 1900
        try:
            return date(year, months[month.casefold()], int(day)).isoformat()
        except (KeyError, ValueError):
            return None
    formats = ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d"]
    if TEXT_DATE_ORDER == "DMY":
        formats += ["%d/%m/%Y", "%d/%m/%y"]
    else:
        formats += ["%m/%d/%Y", "%m/%d/%y"]
    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    return None


def _find_header_row(ws) -> tuple[int, dict]:
    for index, row in enumerate(ws.iter_rows(min_row=1, max_row=HEADER_SCAN_ROWS, values_only=True), 1):
        values = [label_text(v) for v in row]
        if "date" in values and "hours" in values:
            if values.count("date") > 1 or values.count("hours") > 1:
                raise ValueError("发现重复 Date/Hours 标题，无法确定明细列")
            return index, {v: i+1 for i, v in enumerate(values) if v}
    raise ValueError(f"前 {HEADER_SCAN_ROWS} 行未找到 Date / Hours 表头")


def _get_col(col_map: dict, *candidates) -> int | None:
    """从多个候选列名里找第一个存在的列索引"""
    for name in candidates:
        if name in col_map:
            return col_map[name]
    return None


def extract_sheet(ws, staff_id: str, emp_name: str, source_path="") -> list[dict]:
    header_row, col_map = _find_header_row(ws)
    fields = {
        "Date": _get_col(col_map, "date"),
        "Hours": _get_col(col_map, "hours"),
        "Role/Type": _get_col(col_map, "role / type", "role/type"),
        "Task Nature": _get_col(col_map, "task nature"),
        "Module": _get_col(col_map, "module"),
        "Ref ID": _get_col(col_map, "ref id", "ref id (no need)", "ref/ticket #"),
    }
    if fields["Ref ID"] is None:
        fields["Ref ID"] = next((i for title, i in col_map.items() if title.startswith("ref id (")), None)
    for title, column in fields.items():
        if column is None:
            audit(source_path, ws.title, header_row, "缺少可选列", title)
    stats = Counter()
    records = []
    for row_index, row in enumerate(ws.iter_rows(min_row=header_row+1, values_only=True), header_row+1):
        def g(column):
            value = row[column-1] if column and column <= len(row) else None
            return (value.strip() or None) if isinstance(value, str) else value
        values = {key: g(column) for key, column in fields.items()}
        raw_date, raw_hours = values["Date"], values["Hours"]
        if all(v is None for v in values.values()):
            stats["空行"] += 1
            continue
        if any(label_text(v) in {"total", "subtotal", "grand total", "合计", "总计", "小计"}
               for v in row if isinstance(v, str)):
            stats["合计行"] += 1
            continue
        if label_text(raw_date) == "date" and label_text(raw_hours) == "hours":
            stats["重复标题"] += 1
            continue
        parsed = _parse_date(raw_date, getattr(ws.parent, "epoch", None))
        if parsed is None:
            stats["日期异常"] += 1
            audit(source_path, ws.title, row_index, "日期异常", repr(raw_date))
            continue
        if not FILTER_START_DATE <= parsed <= FILTER_END_DATE:
            stats["范围外"] += 1
            continue
        try:
            if isinstance(raw_hours, bool) or raw_hours is None:
                raise ValueError("空值或布尔值")
            hours = float(raw_hours)
            if not math.isfinite(hours):
                raise ValueError("非有限数字")
        except (TypeError, ValueError, OverflowError):
            stats["工时异常"] += 1
            audit(source_path, ws.title, row_index, "工时异常", repr(raw_hours))
            continue
        if hours < 0 or hours > 24:
            audit(source_path, ws.title, row_index, "工时需核对", f"{hours}（保留，未自动删除）")
        values.update({"Date": parsed, "Hours": hours, "Staff ID": staff_id,
                       "Name": emp_name, "_source": str(source_path),
                       "_sheet": ws.title, "_row": row_index})
        records.append(values)
    logger.info(f"  [明细统计] {ws.title}：保留 {len(records)}；" +
                "；".join(f"{k} {v}" for k, v in stats.items()))
    return records


# ─────────────────────────────────────────────
# Step 4：解析单个源文件
# ─────────────────────────────────────────────
def parse_source(path: Path) -> list[dict]:
    """
    读取源文件，返回标准化记录列表。

    顺序：active → Timesheet → 其他可见工作表；每页仅尝试一次。
    排除示例工号，返回首个含范围内有效明细的页，不合并多个页。
    """
    wb = openpyxl.load_workbook(
        path,
        data_only=True,
        read_only=True,
        keep_links=False,
    )

    try:
        # 候选顺序：active → Timesheet → 其他工作表
        candidates = []
        seen_titles = set()
        active_ws = wb.active
        def add_candidate(ws):
            if ws.title not in seen_titles:
                candidates.append(ws)
                seen_titles.add(ws.title)

        if active_ws is not None:
            add_candidate(active_ws)
        # 优先添加 Timesheet，兼容大小写和首尾空格
        for ws in wb.worksheets:
            if ws.title.strip().lower() == "timesheet":
                add_candidate(ws)

        # 最后添加其余工作表，避免数据页名称不同而漏读
        for ws in wb.worksheets:
            add_candidate(ws)

        if not candidates:
            logger.warning(f"  [跳过] {path.name}：未找到可用 sheet")
            return []

        # 逐个尝试候选页，读取到有效记录后立即返回
        for index, ws in enumerate(candidates):
            if index > 0:
                logger.info(
                    f"  [回退] 上一个候选页无有效数据，"
                    f"继续尝试 sheet：'{ws.title}'"
                )

            source = "active" if ws is active_ws else "备选"
            logger.info(f"  [候选] 使用 {source} sheet：'{ws.title}'")

            # 隐藏页不参与读取，继续尝试下一个候选页
            if getattr(ws, "sheet_state", "visible") != "visible":
                logger.warning(
                    f"  [候选无效] sheet '{ws.title}'：隐藏 sheet，忽略"
                )
                continue

            # 前 10 行扫描 Staff ID / Name，允许不在同一行
            staff_id, emp_name = "", ""

            for scan_row in ws.iter_rows(
                min_row=1,
                max_row=PERSON_SCAN_ROWS,
                values_only=True,
            ):
                row_vals = list(scan_row)

                for i, val in enumerate(row_vals):
                    label = label_text(val)

                    if (
                        not staff_id
                        and label == "staff id"
                        and i + 1 < len(row_vals)
                    ):
                        staff_id = normalize_staff_id(row_vals[i + 1])

                    if (
                        not emp_name
                        and label == "name"
                        and i + 1 < len(row_vals)
                    ):
                        emp_name = str(row_vals[i + 1] or "").strip()

                if staff_id and emp_name:
                    break

            # 示例人员所在页整页跳过，继续检查下一个候选 sheet
            if staff_id.strip().upper() in EXCLUDED_STAFF_IDS:
                logger.info(
                    f"  [排除示例页] {path.name} → sheet '{ws.title}'："
                    f"{staff_id} / {emp_name}，忽略该页全部工时"
                )
                continue
            # 工号为花名册关联键，必须存在；姓名允许为空并提示
            if not staff_id:
                logger.warning(
                    f"  [候选无效] {path.name} → sheet '{ws.title}'："
                    f"未找到有效 Staff ID"
                )
                continue

            if not emp_name:
                audit(path, ws.title, "", "缺少姓名", staff_id)
            # 提取数据；缺少表头时继续尝试下一个候选页
            try:
                rows = extract_sheet(ws, staff_id, emp_name, path)
            except ValueError as e:
                logger.warning(
                    f"  [候选无效] {path.name} → sheet '{ws.title}'："
                    f"表头缺失或数据结构不符合要求（{e}）"
                )
                continue

            if not rows:
                logger.warning(
                    f"  [空数据] '{ws.title}' "
                    f"({emp_name} / {staff_id})：0 行有效数据"
                )
                continue

            logger.info(
                f"  [读取] '{ws.title}' "
                f"({emp_name} / {staff_id})：{len(rows)} 行有效数据"
            )
            return rows

        logger.warning(
            f"  [跳过] {path.name}：所有候选 sheet 均无有效数据"
        )
        return []

    finally:
        wb.close()

# ─────────────────────────────────────────────
# Step 5：写入输出文件
# ─────────────────────────────────────────────
def write_output(records: list[dict], out_path: Path, staff_dict: dict):
    """
    创建干净的 xlsx，写入固定 12 列数据。
    表头加粗 + 浅蓝底色，首行冻结，列宽固定。
    """
    if len(records) > 1_048_575:
        raise ValueError("明细超过 Excel 单 sheet 行数上限，请缩小范围")
    wb = Workbook()
    ws = wb.active
    ws.title = "Timesheet"
    for index, title in enumerate(OUTPUT_COLS, 1):
        cell = ws.cell(1, index, title)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="BDD7EE")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for index, rec in enumerate(records, 2):
        info = staff_dict.get(rec["Staff ID"], {})
        for col, title in enumerate(OUTPUT_COLS, 1):
            value = info.get(title, "-") if title in STAFF_COL_NAMES[1:] else rec.get(title)
            if isinstance(value, str) and len(value) > 32767:
                raise ValueError(f"文本超过 Excel 单元格长度：{rec.get('_source')} 行 {rec.get('_row')} / {title}")
            cell = ws.cell(index, col, value)
            if isinstance(value, str):
                cell.data_type = "s"  # 原样保存以 = 开头的文本，避免变为公式
            if title == "Hours":
                cell.number_format = "0.00"
        # 按员工工号划分相邻记录，覆盖全部12列
        if index == len(records)+1 or records[index-1]["Staff ID"] != rec["Staff ID"]:
            for cell in ws[index]:
                cell.border = Border(bottom=Side(style="thin"))
    widths = [14, 32, 24, 24, 10, 22, 16, 26, 20, 22, 22, 22]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    out_path.parent.mkdir(parents=True, exist_ok=True)
    handle, tmp_name = tempfile.mkstemp(prefix=".timesheet_", suffix=".xlsx", dir=out_path.parent)
    os.close(handle)
    try:
        wb.save(tmp_name)
        if out_path.exists():
            raise FileExistsError(f"拒绝覆盖已有结果：{out_path}")
        os.replace(tmp_name, out_path)
    finally:
        wb.close()
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)
    logger.info(f"[输出] {out_path}（{len(records)} 行）")


# ─────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────
def resolve_path(value):
    path = Path(value).expanduser()
    return path.resolve() if path.is_absolute() else (BASE_DIR / path).resolve()


def main(argv=None):
    global logger, FILTER_START_DATE, FILTER_END_DATE, TEXT_DATE_ORDER
    parser = argparse.ArgumentParser(description="Timesheet Merger：自动读取 TXT 花名册")
    parser.add_argument("--src", default=DEFAULT_SRC_DIR, help="源Excel目录；相对路径基于脚本目录")
    parser.add_argument("--out", default=DEFAULT_OUT_DIR, help="输出目录")
    parser.add_argument("--staff", default=DEFAULT_STAFF_FILE, help="花名册TXT，默认脚本旁 teaminfoList.txt")
    parser.add_argument("--staff-encoding", default=STAFF_ENCODING, help="TXT编码，默认auto")
    parser.add_argument("--start", default=FILTER_START_DATE, help="开始日期 YYYY-MM-DD（包含）")
    parser.add_argument("--end", default=FILTER_END_DATE, help="结束日期 YYYY-MM-DD（包含）")
    parser.add_argument("--date-order", choices=["DMY", "MDY"], default=TEXT_DATE_ORDER)
    args = parser.parse_args(argv)
    FILTER_START_DATE = date.fromisoformat(args.start).isoformat()
    FILTER_END_DATE = date.fromisoformat(args.end).isoformat()
    TEXT_DATE_ORDER = args.date_order
    if FILTER_START_DATE > FILTER_END_DATE:
        raise ValueError("开始日期不能晚于结束日期")
    src_dir, out_dir, staff_path = map(resolve_path, [args.src, args.out, args.staff])
    if not src_dir.is_dir():
        raise NotADirectoryError(f"源文件目录不存在或不是目录：{src_dir}")
    if src_dir == out_dir:
        raise ValueError("源目录和输出目录不能相同，避免把汇总结果当作源数据")
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    logger = setup_logger(out_dir / f"run_{timestamp}.log")
    AUDIT_EVENTS.clear()
    logger.info(f"源目录：{src_dir}；花名册：{staff_path}；日期：{FILTER_START_DATE} ~ {FILTER_END_DATE}（包含边界）；文本日期：{TEXT_DATE_ORDER}")
    # 花名册缺失/损坏直接停止，避免输出整批未匹配的部门字段
    roster = load_staff_from_txt(staff_path, args.staff_encoding)
    files = discover_files(src_dir)
    records, skipped, errored = [], [], []
    staff_names, seen_rows, staff_sources = {}, {}, {}
    for index, path in enumerate(files, 1):
        logger.info(f"[{index}/{len(files)}] {path.name}")
        try:
            rows = parse_source(path)
        except Exception as exc:
            logger.exception(f"[文件异常] {path.name}：{exc}")
            audit(path, "", "", "文件异常", str(exc))
            errored.append(path.name)
            continue
        if not rows:
            skipped.append(path.name)
            continue
        for rec in rows:
            sid = rec["Staff ID"]
            staff_names.setdefault(sid, set()).add(rec["Name"])
            sources = staff_sources.setdefault(sid, set())
            if sources and str(path) not in sources:
                audit(path, rec["_sheet"], rec["_row"], "同工号多个文件", f"{sid}；此前：{', '.join(sorted(sources))}；数据保留")
            sources.add(str(path))
            key = tuple(rec.get(title) for title in OUTPUT_COLS[:8])
            if key in seen_rows:
                audit(path, rec["_sheet"], rec["_row"], "疑似重复明细", f"首次位置：{seen_rows[key]}；数据保留")
            else:
                seen_rows[key] = f"{path.name}/{rec['_sheet']}/{rec['_row']}"
        records.extend(rows)
    for sid, names in staff_names.items():
        if len(names) > 1:
            audit("汇总", "", "", "同工号姓名不一致", f"{sid}：{sorted(names)}")
        if roster.get(sid) is None:
            audit("汇总", "", "", "工号未匹配花名册", sid)
    # 保持原文件/明细顺序，不自动去重、不自动合并多个有效sheet
    if AUDIT_EVENTS:
        audit_path = out_dir / f"audit_{timestamp}.csv"
        with audit_path.open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=["File", "Sheet", "Row", "Reason", "Detail"])
            writer.writeheader()
            writer.writerows(AUDIT_EVENTS)
        logger.info(f"[核对清单] {audit_path}")
    if records:
        write_output(records, out_dir / f"output_{timestamp}.xlsx", roster)
    else:
        logger.warning("无范围内有效数据，未生成Excel")
    for sid, names in staff_names.items():
        total = sum((Decimal(str(r["Hours"])) for r in records if r["Staff ID"] == sid), Decimal("0"))
        logger.info(f"[员工] {sid} / {' | '.join(sorted(names))} / 总工时 {total}")
    logger.info(f"合计：{len(files)} 文件；{len(records)} 行；{len(staff_names)} 人；跳过 {len(skipped)}；异常 {len(errored)}；核对项 {len(AUDIT_EVENTS)}")
    logger.info(f"跳过文件：{skipped}；异常文件：{errored}")
    return 2 if errored or not records else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        logger.error(f"[终止] {type(exc).__name__}: {exc}")
        sys.exit(1)
