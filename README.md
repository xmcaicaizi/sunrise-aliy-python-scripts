# sunrise-aliy-python-scripts

用于 aliy 的 Python 脚本仓库

## 环境配置

本项目使用 `uv` 作为包管理工具：

```bash
uv sync
```

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
