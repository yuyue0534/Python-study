好，需求清晰了。方案确定如下：

**选方案 B 简化版** — 配置文件 + 主脚本，同事只需改 `config.yaml` 里的输入文件路径，其他参数固定，运行一条命令即可。加简单的控制台进度提示。

## 文件结构

```
project/
├── config.yaml          ← 同事只需改这里（输入文件路径）
├── pivot_runner.py      ← 主脚本，不需要动
└── output/              ← 自动创建，结果放这里
    └── pivot_YYYYMMDD_HHMMSS.xlsx
```

## 模块拆解

```
pivot_runner.py
├── 1. load_config()         读取 config.yaml，校验参数
├── 2. load_data()           读取原始 xlsx，校验必要列是否存在
├── 3. build_computed_cols() 复现 Zero_Bill 计算列
├── 4. apply_filters()       按固定条件筛选（SCHEME_CODE + BILL_PERIOD）
├── 5. build_pivot()         构建透视表（groupby + agg）
└── 6. export()              写出 xlsx，格式对齐透视表样式
```

每一步都有 `print` 进度提示和错误说明，方便同事判断出了什么问题。现在开始写代码：关键发现：`Zero_Bill`、`RS_SUBMIT`、`RS_SUBMIT_MEM`、`PAID`、`PAID_MEM` 这五列在原始数据里都是 **Excel 公式字符串**，读出来是公式文本不是值。这意味着 SQL 导出的新文件里这些列可能不存在，需要在 Python 里完整复现。现在开始写代码：全流程跑通。把输出文件复制到 outputs 给你下载：三个文件：主脚本、配置文件、输出样例。

---

## 使用说明（给同事）

**第一次准备：**
```bash
pip install pandas openpyxl pyyaml numpy
```

**目录结构：**
```
项目文件夹/
├── pivot_runner.py   ← 不动
├── config.yaml       ← 只改这里
├── data/
│   └── export.xlsx   ← 把SQL导出的文件放这里
└── output/           ← 自动生成，结果在这里
```

**每次使用：**
1. 把 SQL 导出的 xlsx 放进 `data/` 文件夹
2. 打开 `config.yaml`，把 `input_file` 改成实际文件名
3. 运行 `python pivot_runner.py`
4. 结果在 `output/` 里，文件名带时间戳

---

## 几点说明

**关于计算列复现**：`Zero_Bill`、`PAID_MEM` 等五列在你现有的 xlsx 里是 Excel 公式，脚本会自动检测——如果 SQL 导出已经是计算好的数值就直接用，如果读到的是公式字符串会按规则重新计算并给出提示。

**关于筛选条件**：`SCHEME_CODE=FD`、`BILL_PERIOD=202511` 固定写在 `config.yaml` 里，同事不需要改代码，改配置即可。如果以后要换期数，改 `config.yaml` 一行就够。

**关于错误提示**：如果文件路径错、Sheet 名字对不上、或 SQL 导出少了字段，脚本会给出中文说明，告诉同事具体问题在哪，不会只报英文报错。
