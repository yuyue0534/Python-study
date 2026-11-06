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
