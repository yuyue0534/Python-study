--- a/format_timesheet_merger.py
+++ b/format_timesheet_merger.py
@@
 Sheet 选取规则：
   1. 优先使用 wb.active（默认打开页）
-  2. active 不可用时，fallback 到 sheetname == 'Timesheet'
+  2. active 不存在、隐藏、缺少人员信息/表头或无有效数据时，尝试 'Timesheet'
   3. 其他 sheet 一律跳过
@@
-    for row_idx in range(1, 16):
-        row_vals = [str(c.value or "").strip().lower() for c in ws[row_idx]]
+    for row_idx, row in enumerate(
+        ws.iter_rows(min_row=1, max_row=15), start=1
+    ):
+        row_vals = [str(c.value or "").strip().lower() for c in row]
@@
 def parse_source(path: Path) -> list[dict]:
     """
     读取源文件，返回标准化记录列表。
-    Sheet 选取：优先 active sheet，fallback 到 'Timesheet'。
+    优先读取 active；无有效记录时尝试 Timesheet。
+    两者为同一工作表时，只读取一次。
     """
     wb = openpyxl.load_workbook(path, data_only=True, read_only=True, keep_links=False)
-
-    # ── 选取目标 sheet ──
-    active_ws = wb.active
-    if active_ws is not None:
-        ws = active_ws
-        logger.info(f"  [候选] 使用 active sheet：'{ws.title}'")
-    else:
-        ws = None
+    try:
+        # ── 构建候选列表：active 优先，Timesheet 兜底 ──
+        candidates = []
+        active_ws = wb.active
+        if active_ws is not None:
+            candidates.append(active_ws)
+
         for sname in wb.sheetnames:
-            if sname.lower() == "timesheet":
+            if sname.strip().lower() == "timesheet":
                 ws = wb[sname]
-                logger.info(f"  [候选] active 不可用，fallback 到 sheet：'{sname}'")
+                if all(ws.title != item.title for item in candidates):
+                    candidates.append(ws)
                 break
-        if ws is None:
-            logger.warning(f"  [跳过] {path.name}：未找到可用 sheet")
-            wb.close()
-            return []
+
+        for index, ws in enumerate(candidates):
+            source = "active" if ws is active_ws else "fallback Timesheet"
+            logger.info(f"  [候选] 使用 {source} sheet：'{ws.title}'")
+            rows = _parse_source_sheet(path, ws)
+            if rows:
+                return rows
+            if index + 1 < len(candidates):
+                logger.info(
+                    f"  [回退] '{ws.title}' 无有效记录，继续尝试 Timesheet"
+                )
+
+        logger.warning(f"  [跳过] {path.name}：所有候选 sheet 均无有效数据")
+        return []
+    finally:
+        wb.close()
+
+
+def _parse_source_sheet(path: Path, ws) -> list[dict]:
+    """解析一个候选 sheet；不可用或无有效数据时返回空列表。"""
 
     # ── 隐藏 sheet 检测 ──
     try:
         if ws.sheet_state != "visible":
-            logger.warning(f"  [跳过] sheet '{ws.title}'：隐藏 sheet，忽略")
-            wb.close()
+            logger.warning(f"  [候选无效] sheet '{ws.title}'：隐藏 sheet，忽略")
             return []
     except AttributeError:
         pass
@@
     if not staff_id and not emp_name:
-        logger.warning(f"  [跳过] {path.name} → sheet '{ws.title}'：未找到 Staff ID / Name")
-        wb.close()
+        logger.warning(
+            f"  [候选无效] {path.name} → sheet '{ws.title}'：未找到 Staff ID / Name"
+        )
         return []
@@
     except ValueError as e:
-        logger.warning(f"  [警告] {path.name} → '{ws.title}'：{e}，已跳过")
+        logger.warning(f"  [候选无效] {path.name} → '{ws.title}'：{e}")
         rows = []
 
-    wb.close()
     return rows


=============================

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
        # 构建候选列表：active 优先，Timesheet 兜底
        candidates = []
        active_ws = wb.active

        if active_ws is not None:
            candidates.append(active_ws)

        for sname in wb.sheetnames:
            if sname.strip().lower() == "timesheet":
                ws = wb[sname]

                if all(ws.title != item.title for item in candidates):
                    candidates.append(ws)

                break

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

            source = "active" if ws is active_ws else "Timesheet"
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

===========================

-        # 构建候选列表：active 优先，Timesheet 兜底
+        # 候选顺序：active → Timesheet → 其他工作表
         candidates = []
+        seen_titles = set()
         active_ws = wb.active
 
+        def add_candidate(ws):
+            if ws.title not in seen_titles:
+                candidates.append(ws)
+                seen_titles.add(ws.title)
+
         if active_ws is not None:
-            candidates.append(active_ws)
+            add_candidate(active_ws)
 
-        for sname in wb.sheetnames:
-            if sname.strip().lower() == "timesheet":
-                ws = wb[sname]
-
-                if all(ws.title != item.title for item in candidates):
-                    candidates.append(ws)
-
-                break
+        # 优先添加 Timesheet，兼容大小写和首尾空格
+        for ws in wb.worksheets:
+            if ws.title.strip().lower() == "timesheet":
+                add_candidate(ws)
+
+        # 最后添加其余工作表，避免数据页名称不同而漏读
+        for ws in wb.worksheets:
+            add_candidate(ws)
@@
-            source = "active" if ws is active_ws else "Timesheet"
+            source = "active" if ws is active_ws else "备选"
             logger.info(f"  [候选] 使用 {source} sheet：'{ws.title}'")


===============================================================


@@
 DEFAULT_OUT_DIR = "./output"
 
+# 示例人员工号：对应 sheet 整页跳过，不提取工时
+EXCLUDED_STAFF_IDS = {"XCS01818"}
+
@@
                 if staff_id and emp_name:
                     break
 
+            # 示例人员所在页整页跳过，继续检查下一个候选 sheet
+            if staff_id.strip().upper() in EXCLUDED_STAFF_IDS:
+                logger.info(
+                    f"  [排除示例页] {path.name} → sheet '{ws.title}'："
+                    f"{staff_id} / {emp_name}，忽略该页全部工时"
+                )
+                continue
+
             # 保留原规则：Staff ID / Name 至少找到一个即可
             if not staff_id and not emp_name:
