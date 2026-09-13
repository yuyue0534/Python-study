#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
count_and_copy_gui.py

功能（与命令行版 count_and_copy.py 完全一致，只是加了图形界面）：
1. 递归遍历指定源文件夹（包含所有子文件夹），统计每个文件的物理总行数
   （含空行、注释，只要是一行就算），并汇总出总行数。
2. 自动跳过图片、压缩包、可执行文件等二进制/非文本类文件（不统计、不复制）。
3. 将所有符合条件的文件"拍平"复制到一个新文件夹中（不保留原目录层级），
   原文件夹结构和内容不做任何改动。
   - 拍平复制时如遇到同名文件冲突，会在文件名前加上原来的相对路径前缀去重，
     例如 moduleA/utils.py -> moduleA_utils.py
        moduleB/utils.py -> moduleB_utils.py
   - 如果加前缀后仍然冲突（极少见），会在文件名末尾追加数字序号，
     确保任何文件都不会被覆盖丢失。
4. 输出一份统计信息 TXT，内容包括每个文件的相对路径 + 对应行数，
   以及最后的总行数汇总、被跳过的文件列表。

依赖：
    pip install PyQt5

运行：
    python count_and_copy_gui.py
"""

import os
import sys
import shutil
import traceback

from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QFileDialog, QVBoxLayout, QHBoxLayout, QTextEdit, QMessageBox,
    QProgressBar, QGroupBox
)


# ------------------------------------------------------------------
# 核心逻辑（与命令行版本一致）
# ------------------------------------------------------------------

SKIP_EXTENSIONS = {
    # 图片
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".ico", ".webp", ".tiff", ".tif",
    ".svg", ".psd", ".ai", ".raw", ".heic", ".heif",
    # 压缩包
    ".zip", ".rar", ".7z", ".tar", ".gz", ".bz2", ".xz", ".tgz", ".iso",
    # 音视频
    ".mp3", ".mp4", ".wav", ".flac", ".avi", ".mov", ".mkv", ".ogg", ".wmv",
    ".m4a", ".flv", ".webm",
    # 可执行/二进制/编译产物
    ".exe", ".dll", ".so", ".dylib", ".bin", ".o", ".obj", ".class", ".pyc",
    ".pyo", ".a", ".lib", ".app", ".msi", ".deb", ".rpm",
    # 字体
    ".ttf", ".otf", ".woff", ".woff2", ".eot",
    # 常见文档二进制格式（非纯文本）
    ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx",
    # 数据库/其他二进制
    ".db", ".sqlite", ".sqlite3", ".dat",
}


def is_binary_file(filepath, chunk_size=8192):
    """简单二进制探测：文件开头是否包含空字节 \\x00"""
    try:
        with open(filepath, "rb") as f:
            chunk = f.read(chunk_size)
        return b"\x00" in chunk
    except Exception:
        return True


def should_skip(filepath):
    """判断文件是否应该跳过（图片/压缩包/二进制等）"""
    ext = os.path.splitext(filepath)[1].lower()
    if ext in SKIP_EXTENSIONS:
        return True
    if is_binary_file(filepath):
        return True
    return False


def count_lines(filepath):
    """统计文件物理总行数（含空行/注释，兼容无末尾换行符的情况）"""
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        return len(lines)
    except Exception:
        return None


def build_flat_name(rel_path):
    """根据相对路径生成拍平后的文件名：moduleA/utils.py -> moduleA_utils.py"""
    rel_path = rel_path.replace("\\", "/")
    parts = rel_path.split("/")
    return "_".join(parts)


def unique_target_path(target_dir, flat_name):
    """如果拍平后文件名仍冲突，追加数字序号，确保不覆盖任何文件"""
    target_path = os.path.join(target_dir, flat_name)
    if not os.path.exists(target_path):
        return target_path
    name, ext = os.path.splitext(flat_name)
    counter = 1
    while True:
        candidate = os.path.join(target_dir, f"{name}_{counter}{ext}")
        if not os.path.exists(candidate):
            return candidate
        counter += 1


def run_count_and_copy(source_dir, target_dir, stats_output, log_callback=None, progress_callback=None):
    """
    执行统计+拍平复制的主逻辑。
    log_callback(str): 用于实时输出日志（可选）
    progress_callback(current, total): 用于更新进度（可选）
    返回: (file_line_counts, skipped_files, total_lines)
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    source_dir = os.path.abspath(source_dir)
    target_dir = os.path.abspath(target_dir)

    if not os.path.isdir(source_dir):
        raise FileNotFoundError(f"源文件夹不存在: {source_dir}")

    os.makedirs(target_dir, exist_ok=True)

    # 先收集所有文件路径，方便计算总数用于进度条
    all_files = []
    for root, dirs, files in os.walk(source_dir):
        dirs[:] = [d for d in dirs if os.path.join(root, d) != target_dir]
        for filename in files:
            all_files.append(os.path.join(root, filename))

    total_count = len(all_files)
    file_line_counts = []
    skipped_files = []
    total_lines = 0

    for idx, filepath in enumerate(all_files, start=1):
        rel_path = os.path.relpath(filepath, source_dir)

        if should_skip(filepath):
            skipped_files.append(rel_path)
            log(f"[跳过] {rel_path}")
        else:
            line_count = count_lines(filepath)
            if line_count is None:
                skipped_files.append(rel_path)
                log(f"[跳过-读取失败] {rel_path}")
            else:
                file_line_counts.append((rel_path, line_count))
                total_lines += line_count

                flat_name = build_flat_name(rel_path)
                target_path = unique_target_path(target_dir, flat_name)
                shutil.copy2(filepath, target_path)
                log(f"[已复制] {rel_path}  ->  {os.path.basename(target_path)}  ({line_count} 行)")

        if progress_callback:
            progress_callback(idx, total_count)

    # 排序 & 写统计TXT
    file_line_counts.sort(key=lambda x: x[0])

    with open(stats_output, "w", encoding="utf-8") as f:
        f.write("文件行数统计报告\n")
        f.write(f"源文件夹: {source_dir}\n")
        f.write(f"拍平复制目标文件夹: {target_dir}\n")
        f.write("=" * 60 + "\n\n")

        f.write(f"{'相对路径':<50} 行数\n")
        f.write("-" * 60 + "\n")
        for rel_path, line_count in file_line_counts:
            f.write(f"{rel_path:<50} {line_count}\n")

        f.write("-" * 60 + "\n")
        f.write(f"文件总数: {len(file_line_counts)}\n")
        f.write(f"总行数: {total_lines}\n")

        if skipped_files:
            f.write("\n" + "=" * 60 + "\n")
            f.write(f"已跳过的文件（图片/压缩包/二进制等），共 {len(skipped_files)} 个:\n")
            f.write("-" * 60 + "\n")
            for rel_path in sorted(skipped_files):
                f.write(f"{rel_path}\n")

    return file_line_counts, skipped_files, total_lines


# ------------------------------------------------------------------
# 后台线程：避免长时间任务阻塞GUI界面
# ------------------------------------------------------------------

class WorkerThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal(list, list, int, str)  # file_line_counts, skipped_files, total_lines, stats_output
    error_signal = pyqtSignal(str)

    def __init__(self, source_dir, target_dir, stats_output):
        super().__init__()
        self.source_dir = source_dir
        self.target_dir = target_dir
        self.stats_output = stats_output

    def run(self):
        try:
            file_line_counts, skipped_files, total_lines = run_count_and_copy(
                self.source_dir,
                self.target_dir,
                self.stats_output,
                log_callback=self.log_signal.emit,
                progress_callback=lambda cur, total: self.progress_signal.emit(cur, total),
            )
            self.finished_signal.emit(file_line_counts, skipped_files, total_lines, self.stats_output)
        except Exception as e:
            self.error_signal.emit(f"{e}\n\n{traceback.format_exc()}")


# ------------------------------------------------------------------
# GUI 主窗口
# ------------------------------------------------------------------

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("文件行数统计 & 拍平复制工具")
        self.resize(760, 560)

        main_layout = QVBoxLayout()

        # --- 路径选择区 ---
        path_group = QGroupBox("路径设置")
        path_layout = QVBoxLayout()

        # 源文件夹
        src_layout = QHBoxLayout()
        src_layout.addWidget(QLabel("源文件夹："))
        self.src_edit = QLineEdit()
        self.src_edit.setPlaceholderText("请选择需要统计和复制的源文件夹")
        src_btn = QPushButton("浏览...")
        src_btn.clicked.connect(self.choose_source_dir)
        src_layout.addWidget(self.src_edit)
        src_layout.addWidget(src_btn)
        path_layout.addLayout(src_layout)

        # 目标文件夹
        tgt_layout = QHBoxLayout()
        tgt_layout.addWidget(QLabel("目标文件夹："))
        self.tgt_edit = QLineEdit()
        self.tgt_edit.setPlaceholderText("拍平复制后的文件存放位置")
        tgt_btn = QPushButton("浏览...")
        tgt_btn.clicked.connect(self.choose_target_dir)
        tgt_layout.addWidget(self.tgt_edit)
        tgt_layout.addWidget(tgt_btn)
        path_layout.addLayout(tgt_layout)

        # 统计TXT路径（可选）
        out_layout = QHBoxLayout()
        out_layout.addWidget(QLabel("统计TXT路径："))
        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText("留空则默认生成在 目标文件夹/stats.txt")
        out_btn = QPushButton("浏览...")
        out_btn.clicked.connect(self.choose_output_file)
        out_layout.addWidget(self.out_edit)
        out_layout.addWidget(out_btn)
        path_layout.addLayout(out_layout)

        path_group.setLayout(path_layout)
        main_layout.addWidget(path_group)

        # --- 操作按钮 ---
        btn_layout = QHBoxLayout()
        self.run_btn = QPushButton("开始统计并复制")
        self.run_btn.setMinimumHeight(36)
        self.run_btn.clicked.connect(self.start_run)
        btn_layout.addWidget(self.run_btn)
        main_layout.addLayout(btn_layout)

        # --- 进度条 ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)

        # --- 日志输出区 ---
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout()
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        log_layout.addWidget(self.log_edit)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)

        # --- 结果摘要 ---
        self.summary_label = QLabel("尚未运行")
        self.summary_label.setStyleSheet("font-weight: bold;")
        main_layout.addWidget(self.summary_label)

        self.setLayout(main_layout)

    # ---------------- 路径选择槽函数 ----------------

    def choose_source_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择源文件夹")
        if path:
            self.src_edit.setText(path)

    def choose_target_dir(self):
        path = QFileDialog.getExistingDirectory(self, "选择目标文件夹")
        if path:
            self.tgt_edit.setText(path)

    def choose_output_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "选择统计TXT保存路径", "stats.txt", "文本文件 (*.txt)")
        if path:
            self.out_edit.setText(path)

    # ---------------- 运行逻辑 ----------------

    def start_run(self):
        source_dir = self.src_edit.text().strip()
        target_dir = self.tgt_edit.text().strip()
        stats_output = self.out_edit.text().strip()

        if not source_dir:
            QMessageBox.warning(self, "提示", "请先选择源文件夹")
            return
        if not os.path.isdir(source_dir):
            QMessageBox.warning(self, "提示", f"源文件夹不存在：\n{source_dir}")
            return
        if not target_dir:
            QMessageBox.warning(self, "提示", "请先选择目标文件夹")
            return

        source_dir_abs = os.path.abspath(source_dir)
        target_dir_abs = os.path.abspath(target_dir)

        # 防止目标文件夹是源文件夹本身或源文件夹的子目录，避免死循环/污染源数据
        try:
            common = os.path.commonpath([source_dir_abs, target_dir_abs])
        except ValueError:
            common = None
        if common == source_dir_abs:
            QMessageBox.warning(
                self, "提示",
                "目标文件夹不能是源文件夹本身或其子文件夹，请选择一个独立的目标文件夹。"
            )
            return

        if not stats_output:
            os.makedirs(target_dir_abs, exist_ok=True)
            stats_output = os.path.join(target_dir_abs, "stats.txt")
        stats_output = os.path.abspath(stats_output)

        # 确认覆盖提示（如果目标文件夹已有内容）
        if os.path.isdir(target_dir_abs) and os.listdir(target_dir_abs):
            reply = QMessageBox.question(
                self, "确认",
                f"目标文件夹已存在内容：\n{target_dir_abs}\n\n"
                "程序不会清空该文件夹，但如果拍平后文件名重复，会自动加序号避免覆盖。\n是否继续？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        self.log_edit.clear()
        self.progress_bar.setValue(0)
        self.summary_label.setText("正在运行...")
        self.run_btn.setEnabled(False)

        self.worker = WorkerThread(source_dir_abs, target_dir_abs, stats_output)
        self.worker.log_signal.connect(self.append_log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.start()

    def append_log(self, msg):
        self.log_edit.append(msg)

    def update_progress(self, current, total):
        if total > 0:
            percent = int(current / total * 100)
            self.progress_bar.setValue(percent)

    def on_finished(self, file_line_counts, skipped_files, total_lines, stats_output):
        self.run_btn.setEnabled(True)
        self.progress_bar.setValue(100)
        summary = (
            f"完成！统计文件数: {len(file_line_counts)}  |  "
            f"总行数: {total_lines}  |  跳过文件数: {len(skipped_files)}\n"
            f"统计报告已保存至: {stats_output}"
        )
        self.summary_label.setText(summary)
        self.append_log("\n===== 运行完成 =====")
        self.append_log(summary)
        QMessageBox.information(self, "完成", summary)

    def on_error(self, error_msg):
        self.run_btn.setEnabled(True)
        self.summary_label.setText("运行出错")
        self.append_log(f"[错误] {error_msg}")
        QMessageBox.critical(self, "出错了", error_msg)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
