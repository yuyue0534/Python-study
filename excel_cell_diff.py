#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel 按位置逐格比较（Python 3.9+）。
安装：python -m pip install openpyxl
启动：python excel_cell_diff.py
自检：python excel_cell_diff.py --self-test

默认读取 Timesheet，也可修改工作表名。支持 .xlsx/.xlsm，不支持旧版 .xls。
比较单元格存储内容，忽略字体、颜色、数字显示格式及合并区域样式。
公式按公式文本比较，不计算公式；合并单元格仅左上角有实际内容。
空值与空字符串相同；数字 8 与 8.0 相同；文本 00123 与 123 不同。
默认保留空格、换行及大小写；可勾选忽略文本首尾空格。
输入文件只读使用。导出报告含摘要、完整差异明细、按行汇总。
"""
import argparse
import math
import os
import queue
import tempfile
import threading
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit('缺少依赖，请先运行：python -m pip install openpyxl')


class Cancelled(Exception):
    pass


def check_cancel(event):
    if event and event.is_set():
        raise Cancelled('已取消')


def content(value, strip=False):
    """转为用于比较和报告的文本，不使用 Excel 显示格式。"""
    if value is None:
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, datetime):
        return value.isoformat(sep=' ')
    if isinstance(value, (date, time)):
        return value.isoformat()
    if isinstance(value, timedelta):
        return str(value)
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        return str(int(value))
    # openpyxl 3.1 的数组公式对象具有 text 属性。
    value = getattr(value, 'text', value)
    text = str(value)
    return text.strip() if strip else text


@dataclass(frozen=True)
class Difference:
    row: int
    column: int
    before: str
    after: str

    @property
    def address(self):
        return f'{get_column_letter(self.column)}{self.row}'

    @property
    def kind(self):
        return '新增内容' if self.before == '' else ('清空内容' if self.after == '' else '内容修改')


def read_cells(path, sheet, event=None):
    if Path(path).suffix.lower() not in ('.xlsx', '.xlsm'):
        raise ValueError('仅支持 .xlsx / .xlsm；旧版 .xls 请先另存为 .xlsx。')
    book = openpyxl.load_workbook(path, data_only=False, read_only=False, keep_links=False)
    try:
        if sheet not in book.sheetnames:
            raise ValueError(f'{Path(path).name} 中没有工作表“{sheet}”。已有：{", ".join(book.sheetnames)}')
        ws = book[sheet]
        cells = {}
        # 使用已加载的稀疏单元格集合，避免远处仅有样式的单元格令矩形遍历膨胀。
        # _cells 为 openpyxl 内部属性；此处仅读取，不修改工作簿。
        for index, (position, cell) in enumerate(ws._cells.items()):
            if index % 4096 == 0:
                check_cancel(event)
            if cell.value is not None and cell.value != '':
                cells[position] = content(cell.value)
        check_cancel(event)
        return cells
    finally:
        book.close()


def compare_files(a, b, sheet='Timesheet', strip=False, event=None, progress=None):
    if not sheet.strip():
        raise ValueError('请输入工作表名。')
    progress = progress or (lambda text: None)
    progress('正在读取原始文件 A…')
    left = read_cells(a, sheet, event)
    progress('正在读取修改后文件 B…')
    right = read_cells(b, sheet, event)
    progress('正在按行号、列号比较…')
    changes = []
    for index, (row, col) in enumerate(sorted(left.keys() | right.keys())):
        if index % 4096 == 0:
            check_cancel(event)
        old, new = left.get((row, col), ''), right.get((row, col), '')
        if (old.strip() if strip else old) != (new.strip() if strip else new):
            changes.append(Difference(row, col, old, new))
    return {
        'a': str(Path(a).resolve()), 'b': str(Path(b).resolve()), 'sheet': sheet,
        'strip': strip, 'changes': changes, 'created': datetime.now().isoformat(timespec='seconds'),
        'a_rows': max((r for r, c in left), default=0),
        'b_rows': max((r for r, c in right), default=0),
        'a_cols': max((c for r, c in left), default=0),
        'b_cols': max((c for r, c in right), default=0),
    }


def same_path(a, b):
    if Path(a).resolve() == Path(b).resolve():
        return True
    try:
        return os.path.samefile(a, b)
    except OSError:
        return False


def export_report(result, target):
    target = Path(target).resolve()
    if any(same_path(target, result[key]) for key in ('a', 'b')):
        raise ValueError('报告不能覆盖原始文件 A 或 B，请换一个文件名。')
    if target.suffix.lower() != '.xlsx':
        raise ValueError('报告扩展名必须为 .xlsx。')
    book = openpyxl.Workbook()
    summary = book.active
    summary.title = '比对摘要'
    changes = result['changes']
    counts = Counter(d.kind for d in changes)

    def append(ws, values):
        ws.append(values)
        # 原值和公式均作为纯文本写入，避免报告把它们当成可执行公式。
        for cell in ws[ws.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'

    for row in [
        ('项目', '值'), ('原始文件 A', result['a']), ('修改后文件 B', result['b']),
        ('工作表', result['sheet']), ('比对时间', result['created']),
        ('比较方式', '相同行号、列号；公式按公式文本；忽略所有格式'),
        ('忽略文本首尾空格', '是' if result['strip'] else '否'),
        ('A 内容末行 / 末列', f"{result['a_rows']} / {result['a_cols']}"),
        ('B 内容末行 / 末列', f"{result['b_rows']} / {result['b_cols']}"),
        ('有差异的行数', len({d.row for d in changes})), ('差异单元格数', len(changes)),
        ('新增内容', counts['新增内容']), ('清空内容', counts['清空内容']),
        ('内容修改', counts['内容修改']),
        ('说明', '插入或删除行会引起后续位置差异；不推断新增/删除业务记录。'),
        ('合并单元格', '仅比较实际存值的左上角；合并样式本身不比较。'),
        ('显示格式', '日期使用 ISO 文本；数字不套用格式，例如显示为 001 的数字 1 按 1 比较。'),
    ]:
        append(summary, row)
    colors = {'新增内容': 'E2F0D9', '清空内容': 'FCE4D6', '内容修改': 'FFF2CC'}
    # 超过 Excel 单表行数时自动分页，保留完整明细。
    details = None
    for index, diff in enumerate(changes):
        if index % 1000000 == 0:
            details = book.create_sheet('差异明细' if index == 0 else f'差异明细{index // 1000000 + 1}')
            append(details, ('类型', '行号', '列号', '单元格', '原始值 A', '修改后值 B'))
        append(details, (diff.kind, diff.row, diff.column, diff.address, diff.before, diff.after))
        for cell in details[details.max_row]:
            cell.fill = PatternFill('solid', fgColor=colors[diff.kind])
    if details is None:
        details = book.create_sheet('差异明细')
        append(details, ('类型', '行号', '列号', '单元格', '原始值 A', '修改后值 B'))
    grouped = book.create_sheet('按行汇总')
    append(grouped, ('行号', '差异单元格数', '差异位置（过长时分段）'))
    rows = {}
    for diff in changes:
        rows.setdefault(diff.row, []).append(diff.address)
    for row, addresses in rows.items():
        # 分段避免 Excel 单元格 32767 字符限制。
        for start in range(0, len(addresses), 2000):
            append(grouped, (row, len(addresses), ', '.join(addresses[start:start + 2000])))
    for ws in book:
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color='FFFFFF')
            cell.fill = PatternFill('solid', fgColor='285A83')
        widths = [30, 100] if ws is summary else ([14, 12, 12, 16, 60, 60] if ws.title.startswith('差异明细') else [14, 20, 100])
        for index, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(index)].width = width
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(vertical='top', wrap_text=True)
    # 同目录临时文件 + 原子替换，写入失败不会破坏已有报告。
    fd, temporary = tempfile.mkstemp(suffix='.xlsx', dir=target.parent)
    os.close(fd)
    try:
        book.save(temporary)
        os.replace(temporary, target)
    finally:
        book.close()
        if os.path.exists(temporary):
            os.unlink(temporary)


def launch_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    class App:
        PAGE = 500

        def __init__(self, root):
            self.root = root
            self.result = None
            self.filtered = []
            self.page = 0
            self.busy = False
            self.events = queue.Queue()
            self.cancel = threading.Event()
            root.title('Excel 单元格内容比对')
            root.geometry('1150x760')
            root.minsize(850, 560)
            self.a, self.b = tk.StringVar(), tk.StringVar()
            self.sheet = tk.StringVar(value='Timesheet')
            self.strip = tk.BooleanVar(value=False)
            self.kind = tk.StringVar(value='全部差异')
            self.search = tk.StringVar()
            self.status = tk.StringVar(value='选择文件后点击“开始比对”。')
            self.pagination = tk.StringVar()
            form = ttk.Frame(root, padding=14)
            form.pack(fill='x')
            form.columnconfigure(1, weight=1)
            for row, (label, variable) in enumerate((('原始文件 A', self.a), ('修改后文件 B', self.b))):
                ttk.Label(form, text=label).grid(row=row, column=0, padx=(0, 10), pady=5)
                ttk.Entry(form, textvariable=variable).grid(row=row, column=1, sticky='ew')
                ttk.Button(form, text='选择文件…', command=lambda v=variable: self.choose(v)).grid(row=row, column=2, padx=(10, 0))
            options = ttk.Frame(root, padding=(14, 0))
            options.pack(fill='x')
            ttk.Label(options, text='工作表名').pack(side='left')
            ttk.Entry(options, textvariable=self.sheet, width=22).pack(side='left', padx=8)
            ttk.Checkbutton(options, text='忽略文本首尾空格', variable=self.strip).pack(side='left', padx=12)
            self.compare_button = ttk.Button(options, text='开始比对', command=self.compare)
            self.compare_button.pack(side='left', padx=8)
            self.export_button = ttk.Button(options, text='导出完整 Excel 报告', command=self.export, state='disabled')
            self.export_button.pack(side='left')
            ttk.Label(root, text='按相同行号、列号比较；忽略字体、颜色等格式。公式比较公式文本；插入行会导致后续位置差异。', padding=14).pack(anchor='w')
            filters = ttk.Frame(root, padding=(14, 0))
            filters.pack(fill='x')
            combo = ttk.Combobox(filters, textvariable=self.kind, values=['全部差异', '内容修改', '新增内容', '清空内容'], state='readonly', width=14)
            combo.pack(side='left')
            combo.bind('<<ComboboxSelected>>', lambda e: self.filter())
            ttk.Label(filters, text='搜索位置 / 原值 / 新值').pack(side='left', padx=10)
            entry = ttk.Entry(filters, textvariable=self.search, width=35)
            entry.pack(side='left')
            entry.bind('<Return>', lambda e: self.filter())
            ttk.Button(filters, text='筛选', command=self.filter).pack(side='left', padx=8)
            area = ttk.Frame(root, padding=14)
            area.pack(fill='both', expand=True)
            area.rowconfigure(0, weight=1)
            area.columnconfigure(0, weight=1)
            self.tree = ttk.Treeview(area, columns=('kind', 'row', 'address', 'before', 'after'), show='headings')
            for col, title, width in [('kind', '类型', 95), ('row', '行号', 65), ('address', '单元格', 85), ('before', '原始值 A', 350), ('after', '修改后值 B', 350)]:
                self.tree.heading(col, text=title)
                self.tree.column(col, width=width, minwidth=60, stretch=col in ('before', 'after'))
            self.tree.grid(row=0, column=0, sticky='nsew')
            yscroll = ttk.Scrollbar(area, orient='vertical', command=self.tree.yview)
            yscroll.grid(row=0, column=1, sticky='ns')
            xscroll = ttk.Scrollbar(area, orient='horizontal', command=self.tree.xview)
            xscroll.grid(row=1, column=0, sticky='ew')
            self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
            for kind, color in [('新增内容', '#e2f0d9'), ('清空内容', '#fce4d6'), ('内容修改', '#fff2cc')]:
                self.tree.tag_configure(kind, background=color)
            self.tree.bind('<<TreeviewSelect>>', self.detail)
            nav = ttk.Frame(root, padding=(14, 0))
            nav.pack(fill='x')
            ttk.Button(nav, text='上一页', command=lambda: self.turn(-1)).pack(side='left')
            ttk.Button(nav, text='下一页', command=lambda: self.turn(1)).pack(side='left', padx=8)
            ttk.Label(nav, textvariable=self.pagination).pack(side='left')
            ttk.Label(root, text='选中一条差异可在下方查看完整文本；导出报告包含所有差异，不受筛选及分页影响。', padding=(14, 8)).pack(anchor='w')
            self.detail_text = tk.Text(root, height=6, wrap='word', state='disabled')
            self.detail_text.pack(fill='x', padx=14)
            ttk.Label(root, textvariable=self.status, padding=14).pack(anchor='w')
            root.protocol('WM_DELETE_WINDOW', self.close)
            root.after(120, self.poll)

        def choose(self, variable):
            path = filedialog.askopenfilename(filetypes=[('Excel 文件', '*.xlsx *.xlsm')])
            if path:
                variable.set(path)

        def set_busy(self, busy):
            self.busy = busy
            self.compare_button.configure(state='disabled' if busy else 'normal')
            self.export_button.configure(state='normal' if not busy and self.result is not None else 'disabled')

        def work(self, function, event_name):
            self.set_busy(True)
            def run():
                try:
                    self.events.put((event_name, function()))
                except Cancelled:
                    self.events.put(('cancelled', None))
                except Exception as exc:
                    self.events.put(('error', str(exc)))
            threading.Thread(target=run, daemon=True).start()

        def compare(self):
            a, b, sheet, strip = self.a.get().strip(), self.b.get().strip(), self.sheet.get().strip(), self.strip.get()
            if not a or not b or not sheet:
                messagebox.showwarning('缺少信息', '请选择两份文件，并输入工作表名。')
                return
            self.result = None
            self.filtered = []
            self.page = 0
            self.render()
            self.detail_text.configure(state='normal')
            self.detail_text.delete('1.0', 'end')
            self.detail_text.configure(state='disabled')
            self.work(lambda: compare_files(a, b, sheet, strip, self.cancel, lambda s: self.events.put(('progress', s))), 'compared')

        def filter(self):
            needle, kind = self.search.get().casefold(), self.kind.get()
            self.filtered = [d for d in self.result['changes'] if (kind == '全部差异' or d.kind == kind) and (not needle or needle in (d.address + '\n' + d.before + '\n' + d.after).casefold())] if self.result else []
            self.page = 0
            self.render()

        def render(self):
            self.tree.delete(*self.tree.get_children())
            start = self.page * self.PAGE
            for index, diff in enumerate(self.filtered[start:start + self.PAGE], start):
                def preview(text):
                    return text.replace('\r', '\\r').replace('\n', '\\n').replace('\t', '\\t')[:250]
                self.tree.insert('', 'end', iid=str(index), values=(diff.kind, diff.row, diff.address, preview(diff.before), preview(diff.after)), tags=(diff.kind,))
            pages = max(1, (len(self.filtered) + self.PAGE - 1) // self.PAGE)
            self.pagination.set(f'第 {self.page + 1} / {pages} 页；筛选后 {len(self.filtered)} 条差异')

        def turn(self, delta):
            self.page = max(0, min(self.page + delta, max(0, (len(self.filtered) - 1) // self.PAGE)))
            self.render()

        def detail(self, event=None):
            selection = self.tree.selection()
            if not selection:
                return
            diff = self.filtered[int(selection[0])]
            self.detail_text.configure(state='normal')
            self.detail_text.delete('1.0', 'end')
            self.detail_text.insert('1.0', f'{diff.address} · {diff.kind}\n原始值 A：{diff.before!r}\n修改后值 B：{diff.after!r}')
            self.detail_text.configure(state='disabled')

        def export(self):
            if self.result is None:
                return
            path = filedialog.asksaveasfilename(defaultextension='.xlsx', initialfile='Excel差异报告.xlsx', filetypes=[('Excel 报告', '*.xlsx')])
            if path:
                snapshot = self.result
                self.status.set('正在导出完整报告…')
                self.work(lambda: (export_report(snapshot, path), path)[1], 'exported')

        def poll(self):
            try:
                while True:
                    event, value = self.events.get_nowait()
                    if event == 'progress':
                        self.status.set(value)
                        continue
                    self.set_busy(False)
                    if event == 'compared':
                        self.result = value
                        self.set_busy(False)
                        self.filter()
                        changes = value['changes']
                        self.status.set(f'比对完成：{len({d.row for d in changes})} 行、{len(changes)} 个单元格存在差异。' if changes else '比对完成：两份工作表的内容完全一致。')
                    elif event == 'exported':
                        self.status.set(f'已导出：{value}')
                        messagebox.showinfo('导出完成', f'完整报告已保存：\n{value}')
                    elif event == 'error':
                        self.status.set('操作失败。')
                        messagebox.showerror('操作失败', value)
            except queue.Empty:
                pass
            self.root.after(120, self.poll)

        def close(self):
            if self.busy and not messagebox.askyesno('关闭', '任务正在进行，确定关闭窗口吗？'):
                return
            self.cancel.set()
            self.root.destroy()

    root = tk.Tk()
    App(root)
    root.mainloop()


def self_test():
    from openpyxl.styles import Font
    with tempfile.TemporaryDirectory() as folder:
        a, b, out = (Path(folder) / name for name in ('a.xlsx', 'b.xlsx', 'report.xlsx'))
        book = openpyxl.Workbook()
        ws = book.active
        ws.title = 'Timesheet'
        for row in [('Staff ID', 'Name', 'Hours'), ('00123', 'Alice', 8), ('00234', 'Bob', 7.5)]:
            ws.append(row)
        ws['D2'] = '=SUM(C2:C3)'
        ws['E2'] = ' 原文 '
        ws['F2'] = 'line1\nline2'
        ws.merge_cells('G1:H1')
        ws['G1'] = '合并表头'
        book.save(a)
        ws['A2'].font = Font(color='FF0000', bold=True)
        ws['C2'] = 8.0
        ws['J5000'].font = Font(bold=True)
        book.save(b)
        assert not compare_files(a, b)['changes'], '格式或整数小数造成误报'
        ws['B2'] = 'Alicia'
        ws['C3'] = None
        ws['D2'] = '=SUM(C2:C4)'
        ws['E2'] = '原文'
        ws['F2'] = 'line1\nline3'
        ws['A4'] = '00345'
        book.save(b)
        result = compare_files(a, b)
        assert [d.address for d in result['changes']] == ['B2', 'D2', 'E2', 'F2', 'C3', 'A4']
        assert len(compare_files(a, b, strip=True)['changes']) == 5
        assert result['b_rows'] == 4 and result['b_cols'] == 7
        export_report(result, out)
        report = openpyxl.load_workbook(out, data_only=False)
        assert report.sheetnames == ['比对摘要', '差异明细', '按行汇总']
        assert report['差异明细']['E3'].value == '=SUM(C2:C3)'
        assert report['差异明细']['E3'].data_type == 's'
        assert report['差异明细'].max_row == 7
        report.close()
        original = a.read_bytes()
        try:
            export_report(result, a)
            raise AssertionError('未保护输入文件')
        except ValueError:
            pass
        assert a.read_bytes() == original
        try:
            compare_files(a, b, '不存在')
            raise AssertionError('未检查工作表')
        except ValueError:
            pass
        empty = openpyxl.Workbook()
        empty.active.title = 'Timesheet'
        empty.save(b)
        assert all(d.kind == '清空内容' for d in compare_files(a, b)['changes'])
        empty.save(a)
        assert not compare_files(a, b)['changes']
        export_report(compare_files(a, b), out)
        flag = threading.Event()
        flag.set()
        try:
            compare_files(a, b, event=flag)
            raise AssertionError('未响应取消')
        except Cancelled:
            pass
        for wb in (book, empty):
            wb.close()
    print('自检通过：位置比对、格式忽略、数字、空值、空格、换行、公式、合并单元格、空表、报告导出、输入保护及错误处理。')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='按位置逐格比较 Excel 文本内容')
    parser.add_argument('--self-test', action='store_true', help='运行内置自检，无需图形界面')
    args = parser.parse_args()
    if args.self_test:
        self_test()
    else:
        launch_gui()
