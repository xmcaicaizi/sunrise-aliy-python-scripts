# 开发文档

## 环境搭建

```bash
# 安装依赖
uv sync

# 安装开发依赖（包含 pytest）
uv sync --dev
```

## 运行测试

```bash
# 运行所有测试
uv run pytest tests/ -v

# 运行单个测试文件
uv run pytest tests/test_scripts.py -v

# 运行特定测试类
uv run pytest tests/test_scripts.py::TestDetectHeader -v
```

## 项目结构

```
├── scripts/                # 脚本目录
│   ├── read_excel_head.py      # 读取 Excel 前 N 行
│   ├── detect_header.py        # 自动检测表头行
│   ├── validate_columns.py     # 校验列名模板
│   ├── analyze_excel_columns.py # 分析列唯一值
│   ├── filter_excel.py         # 按条件剔除行
│   ├── clean_attendance.py     # 考勤数据清洗
│   ├── split_excel.py          # 按列拆分文件
│   ├── join_excel.py           # 关联两个 Excel
│   ├── summary_by_employee.py  # 按工号汇总
│   ├── summary_by_group.py     # 按维度分组汇总
│   └── abnormal_report.py      # 异常考勤报告
├── tests/                  # 测试目录
│   ├── test_scripts.py         # 基础脚本测试
│   └── test_advanced_scripts.py # 高级脚本测试
├── examples/               # 示例数据（git 忽略）
├── docs/                   # 文档目录
└── pyproject.toml          # 项目配置
```

## 添加新脚本

1. 在 `scripts/` 目录下创建新脚本
2. 遵循现有脚本的模式：
   - 提供函数接口（供其他脚本调用）
   - 提供命令行接口（`main()` 函数）
   - 支持 `--header-row` 和 `-s/--sheet` 参数
   - 自动检测表头行（调用 `detect_header_row`）
3. 在 `tests/` 目录下添加对应测试
4. 更新 `README.md` 添加使用说明

## 脚本开发规范

### 参数命名约定

- `--header-row`: 表头所在行（从 0 开始）
- `-s, --sheet`: 工作表名称或索引
- `-o, --output`: 输出文件路径
- `-c, --column(s)`: 列名

### 错误处理

- 文件不存在时抛出 `FileNotFoundError`
- 列名不存在时抛出 `ValueError` 并提示可用列名
- 使用中文错误信息

### 输出规范

- 打印处理进度和统计信息
- 保存文件后打印保存路径

## 提交前检查

```bash
# 运行测试
uv run pytest tests/ -v

# 确保所有测试通过后再提交
git add .
git commit -m "feat: 描述你的更改"
```
