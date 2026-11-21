"""
Timesheet Merger
================
从多个源文件（员工姓名.xlsx）提取考勤数据，
统一汇总输出为单一 output_YYYYMMDD_HHMM.xlsx。

输出固定 12 列：
  Date | Role/Type | Task Nature | Module | Hours | Ref ID | Staff ID | Name
  | Function | Workstream | Subteam | Sub-Subteam（后 4 列来自 teaminfoList.txt）

使用方法：
  python timesheet_merger.py
  python timesheet_merger.py --src ./源文件目录 --out ./输出目录

Sheet 选取规则：
  1. 优先使用 wb.active（默认打开页）
  2. 无有效数据时，依次回退到 'Timesheet'、其余可见 sheet

Staff ID / Name 扫描：
  - 在前 10 行内扫描，允许不在同一行
"""

import argparse
import logging
import warnings
from datetime import datetime, date, timedelta
from pathlib import Path

import csv
import io
import pandas as pd

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

red_fill = PatternFill("solid", fgColor="FF4C4C")
thick_side = Side(style="thin")
thick_bottom = Border(bottom=thick_side)

warnings.filterwarnings("ignore", category=UserWarning)


# ─────────────────────────────────────────────
# 配置区（按需修改）
# ─────────────────────────────────────────────
DEFAULT_SRC_DIR = "./2026Nov2"
DEFAULT_OUT_DIR = "./output"
# 花名册文件：与脚本位于同一目录，tab 分隔，含标题行（A~L 列）
STAFF_FILE = Path(__file__).resolve().parent / "teaminfoList.txt"

# 输出固定列定义（顺序即为输出顺序）
# OUTPUT_COLS = ["Date", "Role/Type", "Task Nature", "Module", "Hours", "Ref ID", "Staff ID", "Name"]
# 示例人员工号：对应 sheet 整页跳过，不提取工时
EXCLUDED_STAFF_IDS = {"XCS01818"}
# 仅保留 Date 在此范围内的明细，包含开始和结束日期
FILTER_START_DATE = "2026-11-09"
FILTER_END_DATE = "2026-11-15"

# 校验配置，并统一为 YYYY-MM-DD 格式
FILTER_START_DATE = date.fromisoformat(FILTER_START_DATE).isoformat()
FILTER_END_DATE = date.fromisoformat(FILTER_END_DATE).isoformat()

if FILTER_START_DATE > FILTER_END_DATE:
    raise ValueError("日期过滤配置错误：开始日期不能晚于结束日期")

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
    """
    创建同时写控制台和文件的 logger。
    日志文件与输出文件同目录，文件名含时间戳。
    """
    logger = logging.getLogger("timesheet")  # 由 main() 中 setup_logger() 初始化
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    # 控制台 handler
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    # 文件 handler
    fh = logging.FileHandler(log_path, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    return logger

# ─────────────────────────────────────────────
# Staff 花名册：从同目录 teaminfoList.txt 读取
# ─────────────────────────────────────────────
def load_staff_from_file(path: Path = STAFF_FILE) -> dict:
    try:
        # 兼容 UTF-8（含 BOM）与 GBK 编码
        try:
            raw = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError:
            raw = path.read_text(encoding="gbk")

        df = pd.read_csv(
            io.StringIO(raw),
            sep="\t",
            engine="python",
            quoting=csv.QUOTE_NONE,
            on_bad_lines="skip",
            header=None,
            dtype=str,
        )

        df = df.iloc[:, STAFF_COL_POSITIONS]
        df.columns = STAFF_COL_NAMES
        df.dropna(how="all", inplace=True)
        df[STAFF_ID_COL] = df[STAFF_ID_COL].str.strip()
        df = df[df[STAFF_ID_COL].notna() & (df[STAFF_ID_COL] != "") & (df[STAFF_ID_COL] != "SID")]

        result = {}
        for _, row in df.iterrows():
            sid = row[STAFF_ID_COL]
            info = {
                "Function":    str(row.get("Function",    "") or "").strip() or "-",
                "Workstream":  str(row.get("Workstream",  "") or "").strip() or "-",
                "Subteam":     str(row.get("Subteam",     "") or "").strip() or "-",
                "Sub-Subteam": str(row.get("Sub-Subteam", "") or "").strip() or "-",
            }
            result[sid] = info
            result[sid.lstrip("0")] = info

        logger.info(f"[花名册] 已从 {path} 加载 {len(result)} 条记录")
        return result

    except Exception as e:
        logger.warning(f"[花名册] 读取 {path} 失败：{e}，相关列将全部填 '-'")
        return {}


# ─────────────────────────────────────────────
# Step 1：扫描源文件目录
# ─────────────────────────────────────────────
def discover_files(src_dir: Path) -> list[Path]:
    """返回目录下所有 .xlsx 文件路径（排除临时文件）"""
    files = [
        f for f in src_dir.glob("*.xlsx")
        if not f.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"在 {src_dir} 未找到任何 .xlsx 文件")
    logger.info(f"[发现] 共找到 {len(files)} 个源文件")
    return sorted(files)


# ─────────────────────────────────────────────
# Step 2：日期处理工具
# ─────────────────────────────────────────────
_TEXT_DATE_FMTS = [
    "%d-%b-%y",    # 6-Apr-26
    "%d-%b-%Y",    # 6-Apr-2026
    "%d/%m/%Y",    # 06/04/2026
    "%d/%m/%y",    # 06/04/26
    "%m/%d/%Y",    # 04/06/2026 美式
    "%Y-%m-%d",    # 2026-04-06 标准格式兜底
]

def _parse_date(val) -> str | None:
    """
    将各种来源的日期值统一转为 yyyy-mm-dd 字符串。
    返回 None 表示无效日期（如 'Total'），调用方应过滤该行。
    """
    if val is None:
        return None
    if isinstance(val, (datetime, date)):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=int(val))).strftime("%Y-%m-%d")
        except Exception:
            return None
    text = str(val).strip()
    if not text:
        return None
    if "total" in text.lower():
        return None
    for fmt in _TEXT_DATE_FMTS:
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


# ─────────────────────────────────────────────
# Step 3：从单个 sheet 提取数据
# ─────────────────────────────────────────────
def _find_header_row(ws) -> tuple[int, dict]:
    """
    在前 15 行里找含 'date' 和 'hours' 的表头行。
    返回 (行号, {列名小写: 列索引 1-based})
    """
    for row_idx in range(1, 16):
        row_vals = [str(c.value or "").strip().lower() for c in ws[row_idx]]
        if "date" in row_vals and "hours" in row_vals:
            col_map = {v: i + 1 for i, v in enumerate(row_vals) if v}
            return row_idx, col_map
    raise ValueError(f"在 sheet '{ws.title}' 的前 15 行未找到含 'Date' 和 'Hours' 的表头")


def _get_col(col_map: dict, *candidates) -> int | None:
    """从多个候选列名里找第一个存在的列索引"""
    for name in candidates:
        if name in col_map:
            return col_map[name]
    return None


def extract_sheet(ws, staff_id: str, emp_name: str) -> list[dict]:
    """
    从 sheet 提取有效行，只保留 8 个目标字段。
    有效行条件：Date 可解析为日期 且 Hours 为数字。
    """
    header_row, col_map = _find_header_row(ws)
    data_start = header_row + 1

    c_date  = _get_col(col_map, "date")
    c_role  = _get_col(col_map,
                        "role / type", "role/type", "role/type ", "role / type ")
    c_task  = _get_col(col_map,
                        "task nature", "task nature ")
    c_mod   = _get_col(col_map,
                        "module", "module ")
    c_hours = _get_col(col_map,
                        "hours", "hours ")
    c_refid = _get_col(col_map,
                        "ref id", "ref id (no need)", "ref/ticket #",
                        "ref id (mandatory for cr and ce, optional for production incident and ec ticket for now)",
                        "ref id \n(mandatory for cr and ce, optional for production incident and ec ticket for now)",
                        "ref id (no need) ")

    rows = []
    for row in ws.iter_rows(min_row=data_start, values_only=True):

        def g(col_idx):
            if col_idx is None:
                return None
            v = row[col_idx - 1] if col_idx <= len(row) else None
            if isinstance(v, str):
                v = v.strip()
                return None if v == "" else v
            return v

        raw_date  = g(c_date)
        raw_hours = g(c_hours)

        if raw_date is None or raw_hours is None:
            continue
        try:
            hours_val = float(raw_hours)
        except (TypeError, ValueError):
            continue
        parsed_date = _parse_date(raw_date)
        if parsed_date is None:
            continue
        # _parse_date 返回 YYYY-MM-DD，可直接比较日期先后
        if not (FILTER_START_DATE <= parsed_date <= FILTER_END_DATE):
            continue

        rows.append({
            "Date":        parsed_date,
            "Role/Type":   g(c_role),
            "Task Nature": g(c_task),
            "Module":      g(c_mod),
            "Hours":       hours_val,
            "Ref ID":      g(c_refid),
            "Staff ID":    staff_id,
            "Name":        emp_name,
        })
    return rows


# ─────────────────────────────────────────────
# Step 4：解析单个源文件
# ─────────────────────────────────────────────
def parse_source(path: Path) -> list[dict]:
    """
    读取源文件，返回标准化记录列表。

    Sheet 选取：
    1. 优先尝试 active sheet。
    2. active 无有效数据时，继续尝试 Timesheet。
    3. 两者为同一工作表时，只读取一次。
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
                max_row=10,
                values_only=True,
            ):
                row_vals = list(scan_row)

                for i, val in enumerate(row_vals):
                    label = str(val or "").strip().lower()

                    if (
                        not staff_id
                        and label == "staff id"
                        and i + 1 < len(row_vals)
                    ):
                        staff_id = str(row_vals[i + 1] or "").strip()

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
            # 保留原规则：Staff ID / Name 至少找到一个即可
            if not staff_id and not emp_name:
                logger.warning(
                    f"  [候选无效] {path.name} → sheet '{ws.title}'："
                    f"未找到 Staff ID / Name"
                )
                continue

            # 提取数据；缺少表头时继续尝试下一个候选页
            try:
                rows = extract_sheet(ws, staff_id, emp_name)
            except (ValueError, IndexError) as e:
                # 原 _find_header_row 在较短的只读工作表上，
                # 访问超出实际行数的行时可能抛出 IndexError。
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
# 改后
def write_output(records: list[dict], out_path: Path, staff_dict: dict):
    """
    创建干净的 xlsx，写入固定 8 列数据。
    表头加粗 + 浅蓝底色，首行冻结，列宽固定。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Timesheet"

    header_font  = Font(bold=True)
    header_fill  = PatternFill("solid", fgColor="BDD7EE")
    header_align = Alignment(horizontal="center", vertical="center")

    # 写表头
    for col_idx, col_name in enumerate(OUTPUT_COLS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font      = header_font
        cell.fill      = header_fill
        cell.alignment = header_align

    # 写数据
    # 改后
    for row_idx, rec in enumerate(records, 2):
        sid        = rec.get("Staff ID", "")
        staff_info = staff_dict.get(sid, {})
        ws.cell(row_idx,  1).value = rec["Date"]
        ws.cell(row_idx,  2).value = rec["Role/Type"]
        ws.cell(row_idx,  3).value = rec["Task Nature"]
        ws.cell(row_idx,  4).value = rec["Module"]
        ws.cell(row_idx,  5).value = rec["Hours"]
        ws.cell(row_idx,  6).value = rec["Ref ID"]
        ws.cell(row_idx,  7).value = rec["Staff ID"]
        ws.cell(row_idx,  8).value = rec["Name"]
        ws.cell(row_idx,  9).value = staff_info.get("Function",    "-")
        ws.cell(row_idx, 10).value = staff_info.get("Workstream",  "-")
        ws.cell(row_idx, 11).value = staff_info.get("Subteam",     "-")
        ws.cell(row_idx, 12).value = staff_info.get("Sub-Subteam", "-")

    #顺序遍历，相邻行Name不同的时候，在上一行划线，末行补一条线
    def _draw_thick_bottom(r: int):
        for col_idx in range(1, len(OUTPUT_COLS) + 1):
            ws.cell(r, col_idx).border = Border(bottom=thick_side)
    prev_name = ws.cell(2, 8).value
    for row_idx in range(3, len(records) +2):
        cur_name = ws.cell(row_idx, 8).value
        if cur_name != prev_name:
            _draw_thick_bottom(row_idx - 1)     #上一行是该组的末尾
            prev_name = cur_name
    _draw_thick_bottom(len(records) + 1)        #最后一行补线

    # 列宽，12 列全部设置
    for col_idx, width in zip(range(1, 13), [14, 32, 24, 24, 8, 14, 14, 24, 16, 16, 16, 16]):
        ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = width

    # 冻结首行
    ws.freeze_panes = "A2"

    wb.save(out_path)
    logger.info(f"[输出] → {out_path}（{len(records)} 行）")


# ─────────────────────────────────────────────
# 主流程
# ─────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Timesheet Merger")
    parser.add_argument("--src", default=DEFAULT_SRC_DIR, help="源文件目录")
    parser.add_argument("--out", default=DEFAULT_OUT_DIR, help="输出目录")
    args = parser.parse_args()

    src_dir = Path(args.src)
    out_dir = Path(args.out)

    if not src_dir.exists():
        raise FileNotFoundError(f"源文件目录不存在：{src_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    # 初始化 logger（日志文件与输出同目录）
    global logger
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    log_path = out_dir / f"run_{timestamp}.log"
    logger = setup_logger(log_path)
    logger.info(f"{'=' * 50}")
    logger.info(f"Timesheet Merger 启动")
    logger.info(f"源文件目录：{src_dir}")
    logger.info(f"输出目录：{out_dir}")
    logger.info(f"日志文件：{log_path}")
    logger.info(f"{'=' * 50}")

    # 加载花名册（同目录 teaminfoList.txt）
    staff_dict = load_staff_from_file()

    # Step 1：扫描
    files = discover_files(src_dir)

    # Step 2-4：逐文件解析，同时记录异常文件
    all_records: list[dict] = []
    all_staff:   dict[tuple, None] = {}
    skipped:     list[str] = []   # 0行或异常的文件
    errored:     list[str] = []   # 抛出异常的文件

    for idx, path in enumerate(files, 1):
        logger.info(f"")
        logger.info(f"[{idx}/{len(files)}] 处理：{path.name}")
        try:
            rows = parse_source(path)
            if rows:
                all_records.extend(rows)
                for r in rows:
                    all_staff.setdefault((r["Staff ID"], r["Name"]), None)
            else:
                skipped.append(path.name)
        except Exception as e:
            logger.error(f"  [错误] {path.name}：{e}，已跳过")
            errored.append(path.name)

    # Step 5：写入
    logger.info(f"")
    if all_records:
        out_path = out_dir / f"output_{timestamp}.xlsx"
        # 改后
        write_output(all_records, out_path, staff_dict)
    else:
        logger.warning("[跳过] 无有效数据，未生成输出文件")

    # Step 6：统计汇总
    sep = "─" * 50
    logger.info(f"")
    logger.info(sep)
    logger.info("📊 员工统计汇总")
    logger.info(sep)
    logger.info(f"【有效数据】共 {len(all_staff)} 人：")
    for i, (sid, name) in enumerate(all_staff.keys(), 1):
        logger.info(f"  {i:>3}. {sid:<12} {name}")
    if not all_staff:
        logger.info("  （无数据）")

    if skipped:
        logger.info(f"")
        logger.info(f"【无数据/跳过】共 {len(skipped)} 个文件：")
        for i, fname in enumerate(skipped, 1):
            logger.info(f"  {i:>3}. {fname}")

    if errored:
        logger.info(f"")
        logger.info(f"【异常错误】共 {len(errored)} 个文件：")
        for i, fname in enumerate(errored, 1):
            logger.error(f"  {i:>3}. {fname}")

    logger.info(f"")
    logger.info(sep)
    logger.info(
        f"合计：{len(all_records)} 行 / {len(all_staff)} 人 | "
        f"跳过 {len(skipped)} 个 | 异常 {len(errored)} 个"
    )
    logger.info(f"日志已保存至：{log_path}")


if __name__ == "__main__":
    main()
