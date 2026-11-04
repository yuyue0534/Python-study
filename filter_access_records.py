#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""门禁记录 Excel 按工号筛选为 CSV（Python 3.9+）。

安装：python -m pip install openpyxl
交互：python filter_access_records.py
命令：python filter_access_records.py "门禁记录.xlsx" --pins XCS00017 XCS00018
指定输出：python filter_access_records.py "门禁记录.xlsx" --pins "XCS00017,XCS00018" -o "补录.csv"

读取名为“门禁记录”的工作表，第一行为表头。支持 .xlsx/.xlsm；.xls 请另存为 .xlsx。
匹配忽略工号大小写及首尾空白，但输出保留原值与原记录顺序；不去重、不删改原文件。
输出 UTF-8 BOM CSV，固定十列；Excel 日期单元格输出 yyyy-mm-dd HH:MM:SS，文本原样保留。
CSV 不保存 Excel 的颜色、字体等样式。工号/卡号文本保留前导零；数值且格式为纯 000…
时按该格式补零。Excel 已丢失的前导零或超过 15 位数字的精度无法恢复。
公式单元格不自动计算：工号列或已选记录存在公式时停止，请在副本中粘贴为值后重试。
无匹配时不创建 CSV；已有输出文件不覆盖。程序结束会显示每个工号的匹配数量。
"""

import argparse
import csv
import math
from pathlib import Path
import re
import shutil
import sys
import tempfile
from datetime import date, datetime, time
from collections import Counter
from dataclasses import dataclass

HEADERS = ('TTime', 'Device', 'event_point', 'event_no', 'card_no',
           'pin', 'name', 'deptName', 'inorout', 'verify')
SHEET = '门禁记录'

# 每周需要补录的工号，直接在这里修改、增加或删除
DEFAULT_PINS = [
    'XCS00017',
    'XCS00018',
    'XCS00025',
]

class FilterError(Exception):
    """可以直接展示给用户的输入/数据错误。"""


@dataclass
class Result:
    output: object
    scanned: int
    exported: int
    counts: dict
    missing: list


def parse_pins(parts):
    """支持空格、换行、中英文逗号/分号、顿号；去重且保留输入顺序。"""
    if isinstance(parts, str):
        parts = [parts]
    result = []
    for token in re.split(r'[\s,，;；、]+', ' '.join(parts)):
        value = token.strip().strip('\'"“”‘’').strip().upper()
        if value and value not in result:
            result.append(value)
    if not result:
        raise FilterError('请至少输入一个工号。')
    return result


def cell_text(cell, field):
    value = cell.value
    if cell.data_type == 'f':
        raise FilterError(f'{cell.coordinate}（{field}）含公式，请粘贴为值后重试。')
    if cell.data_type == 'e':
        raise FilterError(f'{cell.coordinate}（{field}）含 Excel 错误：{value}')
    if value is None:
        return ''
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        raise FilterError(f'{cell.coordinate}（{field}）为布尔值，请核对原数据。')
    if isinstance(value, datetime):
        return value.isoformat(sep=' ', timespec='microseconds' if value.microsecond else 'seconds')
    if isinstance(value, date):
        return value.isoformat() + (' 00:00:00' if field == 'TTime' else '')
    if isinstance(value, time):
        if field == 'TTime':
            raise FilterError(f'{cell.coordinate} 只有时间、没有日期，请检查 TTime。')
        return value.isoformat()
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            raise FilterError(f'{cell.coordinate}（{field}）不是有效数字。')
        if field == 'TTime':
            raise FilterError(f'{cell.coordinate} 的 TTime 是裸数字，请在 Excel 中设置为日期时间后重试。')
        if field in ('pin', 'card_no'):
            if value != int(value):
                raise FilterError(f'{cell.coordinate}（{field}）包含小数，请核对编号。')
            if abs(value) >= 10 ** 15:
                raise FilterError(f'{cell.coordinate}（{field}）是超过 15 位的数值编号，可能丢失精度；请从原始来源恢复为文本。')
            text = str(int(value))
            if re.fullmatch(r'0+', cell.number_format or '') and value >= 0:
                text = text.zfill(len(cell.number_format))
            return text
        return str(int(value)) if value == int(value) else str(value)
    return str(value)


def filter_excel(source, pins, output=None):
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise FilterError('缺少依赖，请先运行：python -m pip install openpyxl') from exc

    source = Path(source).expanduser().resolve()
    pins = parse_pins(pins)
    if not source.is_file():
        raise FilterError(f'找不到 Excel：{source}')
    if source.suffix.lower() not in ('.xlsx', '.xlsm'):
        raise FilterError('仅支持 .xlsx/.xlsm；旧版 .xls 请在 Excel 中另存为 .xlsx。')
    destination = (Path(output).expanduser().resolve() if output else
                   source.with_name(f'{source.stem}_筛选_{datetime.now():%Y%m%d_%H%M%S_%f}.csv'))
    if destination == source:
        raise FilterError('输出不能覆盖原始 Excel。')
    if destination.suffix.lower() != '.csv':
        raise FilterError('输出文件扩展名必须为 .csv。')
    if not destination.parent.is_dir():
        raise FilterError(f'输出目录不存在：{destination.parent}')
    if destination.exists():
        raise FilterError(f'输出已存在，未覆盖，请换一个文件名：{destination}')

    counts = Counter({pin: 0 for pin in pins})
    scanned = exported = 0
    workbook = load_workbook(source, read_only=True, data_only=False)
    try:
        if SHEET not in workbook.sheetnames:
            raise FilterError(f'找不到工作表“{SHEET}”，实际工作表：{", ".join(workbook.sheetnames)}')
        sheet = workbook[SHEET]
        # 不依赖某些导出程序写错的工作表尺寸，遍历实际 XML 中的记录。
        sheet.reset_dimensions()
        rows = sheet.iter_rows()
        header = next(rows, None)
        if header is None:
            raise FilterError('工作表为空。')
        names = [str(c.value).strip().lstrip('\ufeff') if c.value is not None else '' for c in header]
        missing = [name for name in HEADERS if name not in names]
        repeated = [name for name in HEADERS if names.count(name) > 1]
        if missing or repeated:
            raise FilterError(f'表头不正确。缺少：{missing or "无"}；重复：{repeated or "无"}。第一行必须是表头。')
        unexpected = [name for name in names if name and name not in HEADERS]
        if unexpected:
            raise FilterError(f'存在额外列：{unexpected}。请确认文件只包含指定的十个业务字段。')
        indexes = [names.index(name) for name in HEADERS]
        pin_index = names.index('pin')
        # 先写临时文件，所有数据通过检查后才生成最终 CSV。
        with tempfile.TemporaryDirectory(prefix='access_filter_') as temp_dir:
            temp_csv = Path(temp_dir) / 'result.csv'
            with temp_csv.open('w', encoding='utf-8-sig', newline='') as stream:
                writer = csv.writer(stream, lineterminator='\r\n')
                writer.writerow(HEADERS)
                for excel_row, row in enumerate(rows, start=2):
                    if all(c.value is None or c.value == '' for c in row):
                        continue
                    scanned += 1
                    if pin_index >= len(row):
                        continue
                    raw_pin = cell_text(row[pin_index], 'pin')
                    key = raw_pin.strip().upper()
                    if key not in counts:
                        continue
                    values = [cell_text(row[i], name) if i < len(row) else ''
                              for name, i in zip(HEADERS, indexes)]
                    if not values[0].strip():
                        raise FilterError(f'第 {excel_row} 行工号 {key} 的 TTime 为空，请修正后导出。')
                    writer.writerow(values)
                    counts[key] += 1
                    exported += 1
            if exported:
                # xb 防止误覆盖，也防止运行期间同名文件被其他程序创建。
                created = False
                try:
                    with destination.open('xb') as target:
                        created = True
                        with temp_csv.open('rb') as source_stream:
                            shutil.copyfileobj(source_stream, target)
                except BaseException:
                    if created:
                        destination.unlink(missing_ok=True)
                    raise
    finally:
        workbook.close()
    return Result(destination if exported else None, scanned, exported, dict(counts),
                  [pin for pin in pins if not counts[pin]])


def main(argv=None):
    parser = argparse.ArgumentParser(description='按工号筛选 Excel“门禁记录”并输出 UTF-8 CSV。')
    parser.add_argument('excel', nargs='?', help='原始 .xlsx/.xlsm 文件路径')
    parser.add_argument('--pins', nargs='+', help='可选：覆盖脚本中的 DEFAULT_PINS 工号列表')
    parser.add_argument('-o', '--output', help='输出 CSV 路径；默认保存在 Excel 同目录，自动添加时间戳')
    args = parser.parse_args(argv)
    interactive = not args.excel
    try:
        source = args.excel or input('Excel 文件路径（可拖入终端）：').strip().strip('\'"')
        pins = args.pins or DEFAULT_PINS
        result = filter_excel(source, pins, args.output)
        print(f'\n扫描非空记录：{result.scanned:,} 条')
        for pin, count in result.counts.items():
            print(f'  {pin}: {count:,} 条')
        print(f'共导出：{result.exported:,} 条')
        if result.output:
            print(f'CSV 已保存：{result.output}')
            print('原 Excel 未修改；记录未去重。CSV 可直接用于导入，避免再用 Excel 自动转换编号后保存。')
        else:
            print('没有匹配记录，未创建 CSV。请核对工号和 Excel 文件。')
        if result.missing:
            print('未找到以下工号：' + '、'.join(result.missing))
        code = 0 if result.exported else 2
    except KeyboardInterrupt:
        print('\n已取消。')
        code = 130
    except Exception as exc:
        print(f'\n处理失败：{exc}', file=sys.stderr)
        code = 1
    if interactive and sys.stdin.isatty():
        try:
            input('\n按回车键退出……')
        except (EOFError, KeyboardInterrupt):
            pass
    return code


if __name__ == '__main__':
    sys.exit(main())
