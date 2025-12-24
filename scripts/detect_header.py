"""
自动检测 Excel 多级表头，返回真实表头所在行
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# 考勤表常见的真实表头关键字
HEADER_KEYWORDS = [
    "工号", "部门", "人员类型", "员工状态", "入职日期", "离职日期",
    "日期", "星期", "班次", "考勤组",
    "上班 1 打卡时间", "下班 1 打卡时间",
    "应出勤天数", "实际出勤天数", "迟到次数", "早退次数",
]


def detect_header_row(
    file_path: str,
    keywords: list[str] | None = None,
    max_rows: int = 10,
    sheet_name: str | int = 0,
) -> int:
    """
    自动检测真实表头所在行
    
    Args:
        file_path: Excel 文件路径
        keywords: 用于识别表头的关键字列表，默认使用考勤表关键字
        max_rows: 最多检查的行数，默认 10 行
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        真实表头所在行索引（从 0 开始）
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    if keywords is None:
        keywords = HEADER_KEYWORDS
    
    # 读取前 N 行，不指定 header
    df = pd.read_excel(file_path, header=None, nrows=max_rows, sheet_name=sheet_name)
    
    best_row = 0
    best_match_count = 0
    
    for row_idx in range(len(df)):
        row_values = df.iloc[row_idx].astype(str).tolist()
        match_count = sum(1 for kw in keywords if kw in row_values)
        
        if match_count > best_match_count:
            best_match_count = match_count
            best_row = row_idx
    
    return best_row


def main():
    parser = argparse.ArgumentParser(description="自动检测 Excel 多级表头")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--max-rows", type=int, default=10, help="最多检查的行数，默认 10")
    parser.add_argument("-s", "--sheet", default="0", help="工作表名称或索引，默认 0")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    try:
        header_row = detect_header_row(args.file, max_rows=args.max_rows, sheet_name=sheet)
        print(f"检测到真实表头在第 {header_row + 1} 行（索引 {header_row}）")
        print(f"使用时请设置: --header-row {header_row}")
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
