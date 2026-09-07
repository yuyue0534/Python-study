#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
count_and_copy.py

功能：
1. 递归遍历指定源文件夹（包含所有子文件夹），统计每个文件的物理总行数
   （含空行、注释，只要是一行就算），并汇总出总行数。
2. 自动跳过图片、压缩包、可执行文件等二进制/非文本类文件（不统计、不复制）。
3. 将所有符合条件的文件"拍平"复制到一个新文件夹中（不保留原目录层级），
   原文件夹结构和内容不做任何改动。
   - 拍平复制时如遇到同名文件冲突，会在文件名前加上原来的相对路径前缀去重，
     例如 moduleA/utils.py -> moduleA_utils.py
        moduleB/utils.py -> moduleB_utils.py
4. 输出一份统计信息 TXT，内容包括每个文件的相对路径 + 对应行数，
   以及最后的总行数汇总。

用法：
    python count_and_copy.py <源文件夹> <目标文件夹> [-o 统计TXT路径]

示例：
    python count_and_copy.py ./my_project ./flat_copy -o ./stats.txt
"""

import os
import sys
import shutil
import argparse

# ------------------------------------------------------------------
# 需要跳过的文件后缀（图片、压缩包、二进制、媒体、字体等常见非文本类型）
# 可根据需要自行增减
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
    """
    简单的二进制文件探测：读取文件开头一段内容，
    如果包含空字节(\\x00)，则判定为二进制文件。
    作为后缀名黑名单的兜底判断，避免遗漏未知类型的二进制文件。
    """
    try:
        with open(filepath, "rb") as f:
            chunk = f.read(chunk_size)
        if b"\x00" in chunk:
            return True
        return False
    except Exception:
        # 读取失败（如权限问题等），保守起见当作二进制跳过
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
    """
    统计文件的物理总行数。
    - 含空行、注释行，只要是一行就算。
    - 兼容文件末尾没有换行符的情况（最后一行也会被计入）。
    - 尝试用 utf-8 读取，失败则用 errors='replace' 兜底，避免编码问题导致脚本崩溃。
    """
    try:
        with open(filepath, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()
        # 如果文件完全为空，行数为 0
        return len(lines)
    except Exception as e:
        print(f"[警告] 读取文件失败，跳过统计: {filepath} ({e})")
        return None


def build_flat_name(rel_path):
    """
    根据相对路径生成拍平后的文件名。
    例如: moduleA/utils.py -> moduleA_utils.py
          utils.py（根目录下）-> utils.py
    """
    rel_path = rel_path.replace("\\", "/")
    parts = rel_path.split("/")
    return "_".join(parts)


def unique_target_path(target_dir, flat_name):
    """
    如果拍平后的文件名仍然冲突（极少见，比如原本就存在下划线导致重名），
    则在文件名（不含扩展名）后追加数字序号，确保最终不覆盖任何文件。
    """
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


def main():
    parser = argparse.ArgumentParser(
        description="统计文件夹下所有文件的行数，并将文件拍平复制到新文件夹，同时输出统计TXT"
    )
    parser.add_argument("source", help="源文件夹路径")
    parser.add_argument("target", help="目标文件夹路径（拍平复制的目的地）")
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="统计信息TXT文件路径（默认在目标文件夹下生成 stats.txt）"
    )
    args = parser.parse_args()

    source_dir = os.path.abspath(args.source)
    target_dir = os.path.abspath(args.target)

    if not os.path.isdir(source_dir):
        print(f"[错误] 源文件夹不存在: {source_dir}")
        sys.exit(1)

    # 目标文件夹如果不存在则创建；如果已存在，不清空、不删除已有内容
    os.makedirs(target_dir, exist_ok=True)

    # 统计TXT默认路径
    stats_output = args.output
    if stats_output is None:
        stats_output = os.path.join(target_dir, "stats.txt")
    stats_output = os.path.abspath(stats_output)

    file_line_counts = []  # [(relative_path, line_count), ...]
    skipped_files = []     # 被跳过的文件（图片/压缩包/二进制等）
    total_lines = 0

    # 递归遍历源文件夹
    for root, dirs, files in os.walk(source_dir):
        # 避免把目标文件夹（如果恰好在源文件夹内部）也遍历进去，防止无限自我复制
        dirs[:] = [d for d in dirs if os.path.join(root, d) != target_dir]

        for filename in files:
            filepath = os.path.join(root, filename)
            rel_path = os.path.relpath(filepath, source_dir)

            if should_skip(filepath):
                skipped_files.append(rel_path)
                continue

            line_count = count_lines(filepath)
            if line_count is None:
                # 读取失败，也算作跳过
                skipped_files.append(rel_path)
                continue

            file_line_counts.append((rel_path, line_count))
            total_lines += line_count

            # 拍平复制
            flat_name = build_flat_name(rel_path)
            target_path = unique_target_path(target_dir, flat_name)
            shutil.copy2(filepath, target_path)

    # 按相对路径排序，输出更整齐
    file_line_counts.sort(key=lambda x: x[0])

    # 写统计TXT
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

    # 控制台输出简要结果
    print(f"统计完成，共统计 {len(file_line_counts)} 个文件，总行数: {total_lines}")
    print(f"已跳过 {len(skipped_files)} 个非文本/二进制文件")
    print(f"拍平复制完成，目标文件夹: {target_dir}")
    print(f"统计报告已写入: {stats_output}")


if __name__ == "__main__":
    main()
