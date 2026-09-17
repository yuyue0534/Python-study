#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_timesheet_pdf.py
========================

从可复制文字的 eMPF Weekly Timesheet PDF 中提取 "Weekly Timesheet Summary" 表格，
跨页合并后输出 CSV。

特点：
- 不使用 OCR，直接读取 PDF 内的真实文字层
- 自动寻找真正的 Timesheet 表头
- 能处理第一页同时包含 General Information 的情况
- 自动合并多页表格
- 保留空白单元格
- 默认输出 UTF-8 BOM CSV，可直接用 Excel 打开
- 自动校验 Mon~Sun 合计是否等于 Weekly Total
- 不会覆盖原 PDF

依赖：
    pip install pymupdf

使用：
    python extract_timesheet_pdf.py "input.pdf"

指定输出：
    python extract_timesheet_pdf.py "input.pdf" -o "timesheet.csv"

只提取，不做合计校验：
    python extract_timesheet_pdf.py "input.pdf" --no-validate
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from pathlib import Path
from typing import Iterable

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        print("错误：缺少 PyMuPDF。请先执行：pip install pymupdf", file=sys.stderr)
        sys.exit(1)


# 默认接受类似 XCS00044、ABC12345 这样的员工编号。
# 如果你的其他 PDF 员工编号规则不同，可按实际情况修改。
STAFF_ID_RE = re.compile(r"^[A-Za-z]{2,}\d{4,}$")

# 只要求表头开头这几列固定。
# 后面的列可以动态变化，不再依赖 Weekly Total / Remarks 等固定结尾字段。
FIXED_START_HEADERS = (
    "Staff ID",
    "Staff Name",
)


def clean_cell(value) -> str:
    """把 PDF 单元格文本整理成适合 CSV 的单行字符串。"""
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\u00a0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_header_row(row: list) -> tuple[list[int], list[str]] | None:
    """
    判断一行是否是 Weekly Timesheet 的真正表头。

    PyMuPDF 在第一页有时会把页面上方 General Information 和 Timesheet
    识别成一个大表，因此第一页的原始表格可能出现额外的空列。
    这这里通过过滤空表头列，并检查表头是否以固定字段开头。

    只固定：
        Staff ID
        Staff Name

    后面的字段全部按 PDF 实际表头动态提取。
    """
    normalized = [clean_cell(cell) for cell in row]

    useful_indices = [i for i, value in enumerate(normalized) if value]
    headers = [normalized[i] for i in useful_indices]

    # 至少要有固定的起始字段
    if len(headers) < len(FIXED_START_HEADERS):
        return None

    # 必须以 Staff ID / Staff Name 开头，
    # 但后面的列名和列数允许动态变化
    if tuple(headers[:len(FIXED_START_HEADERS)]) != FIXED_START_HEADERS:
        return None

    return useful_indices, headers


def is_staff_row(row: list[str]) -> bool:
    """判断一行是否像员工数据行。"""
    if not row:
        return False
    staff_id = row[0].strip()
    return bool(STAFF_ID_RE.fullmatch(staff_id))


def find_timesheet_blocks(page) -> list[tuple[list[str], list[list[str]]]]:
    """
    从单页中找到 Timesheet 表格块。

    返回：
        [
            (
                headers,
                [
                    row1,
                    row2,
                    ...
                ]
            )
        ]
    """
    if not hasattr(page, "find_tables"):
        raise RuntimeError(
            "当前 PyMuPDF 版本不支持 page.find_tables()。\n"
            "请升级：pip install -U pymupdf"
        )

    table_finder = page.find_tables()
    blocks: list[tuple[list[str], list[list[str]]]] = []

    for table in table_finder.tables:
        data = table.extract()
        if not data:
            continue

        for row_index, raw_row in enumerate(data):
            header_info = normalize_header_row(raw_row)
            if header_info is None:
                continue

            useful_indices, headers = header_info
            staff_rows: list[list[str]] = []

            for following_row in data[row_index + 1:]:
                values = [
                    clean_cell(
                        following_row[i]
                        if i < len(following_row)
                        else ""
                    )
                    for i in useful_indices
                ]

                if is_staff_row(values):
                    staff_rows.append(values)

            if staff_rows:
                blocks.append((headers, staff_rows))

    return blocks


def extract_timesheet(pdf_path: Path) -> tuple[list[str], list[list[str]]]:
    """跨页提取并合并 Weekly Timesheet Summary。"""
    try:
        doc = pymupdf.open(str(pdf_path))
    except Exception as exc:
        raise RuntimeError(f"无法打开 PDF：{exc}") from exc

    canonical_headers: list[str] | None = None
    all_rows: list[list[str]] = []

    try:
        for page_number, page in enumerate(doc, start=1):
            blocks = find_timesheet_blocks(page)

            if not blocks:
                print(f"[提示] 第 {page_number} 页未找到 Timesheet 数据行。")
                continue

            for headers, rows in blocks:
                if canonical_headers is None:
                    canonical_headers = headers
                elif headers != canonical_headers:
                    raise RuntimeError(
                        f"第 {page_number} 页表头与前页不一致。\n"
                        f"前页：{canonical_headers}\n"
                        f"本页：{headers}"
                    )

                all_rows.extend(rows)

    finally:
        doc.close()

    if canonical_headers is None or not all_rows:
        raise RuntimeError(
            "没有找到可提取的 Weekly Timesheet Summary。\n"
            "请确认 PDF 中存在可复制文字的 Timesheet 表格。"
        )

    return canonical_headers, all_rows


def find_day_column_indices(headers: list[str]) -> list[int]:
    """
    找出 Mon~Sun 七个日期列。

    表头示例：
        Mon 07-Sep
        Tue 08-Sep
        ...
    """
    weekday_prefixes = ("Mon ", "Tue ", "Wed ", "Thu ", "Fri ", "Sat ", "Sun ")
    return [
        i
        for i, header in enumerate(headers)
        if header.startswith(weekday_prefixes)
    ]


def to_number(value: str) -> float | None:
    """把 1 / 0 / 0.5 / 空白转换为数值；非数字返回 None。"""
    value = value.strip()

    if value == "":
        return 0.0

    try:
        return float(value)
    except ValueError:
        return None


def validate_weekly_totals(
    headers: list[str],
    rows: list[list[str]],
) -> list[str]:
    """校验 Mon~Sun 之和是否等于 Weekly Total，返回警告列表。"""
    warnings: list[str] = []

    if "Weekly Total" not in headers:
        return []

    total_index = headers.index("Weekly Total")
    day_indices = find_day_column_indices(headers)

    if len(day_indices) != 7:
        return []

    for row in rows:
        staff_id = row[0] if row else "(unknown)"

        day_values: list[float] = []
        invalid_day_value = False

        for index in day_indices:
            number = to_number(row[index])
            if number is None:
                invalid_day_value = True
                break
            day_values.append(number)

        if invalid_day_value:
            warnings.append(
                f"{staff_id}: Mon~Sun 中存在无法解析的数值。"
            )
            continue

        expected_total = sum(day_values)
        actual_total = to_number(row[total_index])

        if actual_total is None:
            warnings.append(
                f"{staff_id}: Weekly Total 无法解析：{row[total_index]!r}"
            )
            continue

        if abs(expected_total - actual_total) > 1e-9:
            warnings.append(
                f"{staff_id}: Mon~Sun 合计={expected_total:g}，"
                f"Weekly Total={actual_total:g}"
            )

    return warnings


def write_csv(
    output_path: Path,
    headers: list[str],
    rows: Iterable[list[str]],
) -> None:
    """
    写 CSV。

    使用 utf-8-sig（UTF-8 BOM），Windows Excel 直接打开时
    对中文更友好。
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with output_path.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(rows)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "从可复制文字的 eMPF Weekly Timesheet PDF 中提取表格，"
            "跨页合并输出 CSV。"
        )
    )

    parser.add_argument(
        "pdf",
        help="输入 PDF 文件路径",
    )

    parser.add_argument(
        "-o",
        "--output",
        help=(
            "输出 CSV 路径。"
            "默认：<PDF文件名>_timesheet.csv"
        ),
    )

    parser.add_argument(
        "--no-validate",
        action="store_true",
        help="不校验 Mon~Sun 合计与 Weekly Total",
    )

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    pdf_path = Path(args.pdf).expanduser().resolve()

    if not pdf_path.exists():
        print(f"错误：文件不存在：{pdf_path}", file=sys.stderr)
        return 1

    if not pdf_path.is_file():
        print(f"错误：不是文件：{pdf_path}", file=sys.stderr)
        return 1

    if pdf_path.suffix.lower() != ".pdf":
        print("错误：输入文件必须是 PDF。", file=sys.stderr)
        return 1

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = pdf_path.with_name(
            f"{pdf_path.stem}_timesheet.csv"
        )

    # 防止误覆盖原 PDF。
    if output_path == pdf_path:
        print(
            "错误：输出路径不能与输入 PDF 相同。",
            file=sys.stderr,
        )
        return 1

    try:
        headers, rows = extract_timesheet(pdf_path)
    except Exception as exc:
        print(f"提取失败：{exc}", file=sys.stderr)
        return 1

    try:
        write_csv(output_path, headers, rows)
    except Exception as exc:
        print(f"写入 CSV 失败：{exc}", file=sys.stderr)
        return 1

    print()
    print("=" * 68)
    print("Timesheet 提取完成")
    print("=" * 68)
    print(f"输入文件：{pdf_path}")
    print(f"输出文件：{output_path}")
    print(f"提取员工：{len(rows)} 条")
    print(f"字段数量：{len(headers)}")
    print()
    print("字段：")
    for header in headers:
        print(f"  - {header}")

    if not args.no_validate:
        warnings = validate_weekly_totals(headers, rows)

        print()
        if warnings:
            print(f"合计校验：发现 {len(warnings)} 个需要检查的问题")
            for warning in warnings:
                print(f"  [!] {warning}")
        else:
            print("合计校验：通过（Mon~Sun 与 Weekly Total 一致）")

    print()
    print("完成。CSV 可直接用 Excel 打开。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
