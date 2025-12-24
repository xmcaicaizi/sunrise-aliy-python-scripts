# sunrise-aliy-python-scripts

用于 aliy 的 Python 脚本仓库

## 环境配置

本项目使用 `uv` 作为包管理工具：

```bash
uv sync
```

## 使用指南

本工具集面向数据科学家和 HR，用于考勤数据的处理与分析。

### 推荐工作流程

在使用分析类脚本前，请确保数据已经过验证和清洗：

```bash
# 1. 检测表头行（多级表头场景）
uv run python scripts/detect_header.py your_file.xlsx

# 2. 校验列名是否符合考勤表模板
uv run python scripts/validate_columns.py your_file.xlsx --header-row 1

# 3. 清洗数据（剔除周末、非正式员工、离职、无效打卡）
uv run python scripts/clean_attendance.py your_file.xlsx -o cleaned.xlsx

# 4. 在清洗后的数据上进行分析
uv run python scripts/abnormal_report.py cleaned.xlsx -o abnormal.xlsx
uv run python scripts/summary_by_employee.py cleaned.xlsx -o summary.xlsx
```

> ⚠️ 注意：`abnormal_report.py`、`summary_by_employee.py`、`split_excel.py` 等分析脚本应在清洗后的有效考勤数据上运行，否则统计结果可能包含无效记录（如周末、离职员工、无需打卡等）。

## 脚本说明

### scripts/read_excel_head.py

读取 Excel 文件前 N 行，用于判断表头结构（是否为多级表头）。

```bash
# 读取前 5 行（默认）
uv run python scripts/read_excel_head.py examples/test01.xlsx

# 读取前 10 行
uv run python scripts/read_excel_head.py examples/test01.xlsx -n 10
```

### scripts/filter_excel.py

根据指定的列名和值剔除 Excel 数据行。

```bash
# 剔除"星期"列中值为"星期六"、"星期日"的行
uv run python scripts/filter_excel.py examples/test01.xlsx -c "星期" -v "星期六" "星期日" -o output.xlsx

# 指定表头行（多级表头场景，表头在第 2 行，即 --header-row 1）
uv run python scripts/filter_excel.py examples/test01.xlsx -c "人员类型" -v "实习" "外包" --header-row 1 -o output.xlsx
```

参数说明：
- `-c, --column`: 列名
- `-v, --values`: 要剔除的值（可多个）
- `--header-row`: 表头所在行（从 0 开始），默认 0
- `-o, --output`: 输出文件路径

### scripts/analyze_excel_columns.py

分析 Excel 文件，返回每列的唯一值集合。

```bash
# 分析所有列
uv run python scripts/analyze_excel_columns.py examples/test01.xlsx

# 分析指定列
uv run python scripts/analyze_excel_columns.py examples/test01.xlsx -c "星期" "人员类型"

# 以 JSON 格式输出
uv run python scripts/analyze_excel_columns.py examples/test01.xlsx --json
```

参数说明：
- `-c, --columns`: 指定要分析的列名（可多个）
- `--header-row`: 表头所在行（从 0 开始），默认 0
- `--json`: 以 JSON 格式输出

### scripts/detect_header.py

自动检测 Excel 多级表头，返回真实表头所在行。

```bash
uv run python scripts/detect_header.py examples/test01.xlsx
# 输出: 检测到真实表头在第 2 行（索引 1）
```

### scripts/validate_columns.py

校验 Excel 列名是否符合考勤表模板。

```bash
uv run python scripts/validate_columns.py examples/test01.xlsx --header-row 1
```

### scripts/clean_attendance.py

考勤数据清洗一站式脚本，默认规则：
- 剔除周末（星期六、星期日）
- 剔除非正式员工（实习、外包）
- 剔除离职员工
- 剔除无效打卡（缺卡、无需打卡-休息/出差/自由班制/请假）

```bash
# 自动检测表头，应用所有默认规则
uv run python scripts/clean_attendance.py examples/test01.xlsx -o cleaned.xlsx

# 指定表头行
uv run python scripts/clean_attendance.py examples/test01.xlsx --header-row 1 -o cleaned.xlsx

# 保留周末数据
uv run python scripts/clean_attendance.py examples/test01.xlsx --no-weekend -o cleaned.xlsx

# 保留实习/外包
uv run python scripts/clean_attendance.py examples/test01.xlsx --no-intern -o cleaned.xlsx
```

参数说明：
- `--header-row`: 表头所在行（不指定则自动检测）
- `-o, --output`: 输出文件路径
- `--no-weekend`: 不剔除周末
- `--no-intern`: 不剔除实习/外包
- `--no-resigned`: 不剔除离职员工
- `--no-abnormal`: 不剔除无效打卡


### scripts/split_excel.py

按指定列拆分 Excel 文件，每个唯一值生成一个单独的文件。

```bash
# 按部门拆分
uv run python scripts/split_excel.py examples/test01.xlsx -c "部门" -o output_dir

# 按人员拆分
uv run python scripts/split_excel.py examples/test01.xlsx -c "工号"
```

参数说明：
- `-c, --column`: 用于拆分的列名
- `--header-row`: 表头所在行（不指定则自动检测）
- `-o, --output-dir`: 输出目录（不指定则在源文件目录下创建）

### scripts/abnormal_report.py

异常考勤报告生成，筛选缺卡、旷工、迟到、早退等异常记录。

```bash
# 筛选所有异常类型
uv run python scripts/abnormal_report.py examples/test01.xlsx -o abnormal.xlsx

# 只筛选特定异常类型
uv run python scripts/abnormal_report.py examples/test01.xlsx -t 缺卡 旷工 -o abnormal.xlsx
```

支持的异常类型：`缺卡`、`旷工`、`严重迟到`、`迟到`、`早退`

参数说明：
- `-t, --types`: 要筛选的异常类型（可多个）
- `--header-row`: 表头所在行（不指定则自动检测）
- `-o, --output`: 输出文件路径


### scripts/summary_by_employee.py

按工号汇总考勤统计，输出每位员工的出勤天数、迟到次数等汇总数据。

```bash
# 使用默认汇总字段
uv run python scripts/summary_by_employee.py examples/test01.xlsx -o summary.xlsx

# 指定汇总字段
uv run python scripts/summary_by_employee.py examples/test01.xlsx -c "迟到次数" "旷工天数" -o summary.xlsx
```

默认汇总字段：实际出勤天数、迟到次数、严重迟到次数、早退次数、上班缺卡次数、下班缺卡次数、旷工天数、补卡次数

参数说明：
- `-c, --columns`: 要汇总的列名（可多个，不指定则使用默认配置）
- `--header-row`: 表头所在行（不指定则自动检测）
- `-o, --output`: 输出文件路径
