"""
分析 Excel 文件，返回每列的唯一值集合
"""

import argparse
import json
import sys
from pathlib import Path

import pandas as pd


def analyze_excel_columns(
    file_path: str,
    header_row: int = 0,
    columns: list[str] | None = None,
    sheet_name: str | int = 0,
) -> dict[str, set]:
    """
    分析 Excel 文件，返回每列的唯一值集合
    
    Args:
        file_path: Excel 文件路径
        header_row: 表头所在行（从 0 开始），默认第 0 行
        columns: 指定要分析的列名列表，为 None 时分析所有列
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        字典，key 为列名，value 为该列的唯一值 set 集合
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    df = pd.read_excel(file_path, header=header_row, sheet_name=sheet_name)
    
    # 确定要分析的列
    if columns:
        missing = [c for c in columns if c not in df.columns]
        if missing:
            raise ValueError(f"列名不存在: {missing}。可用列名: {list(df.columns)}")
        target_columns = columns
    else:
        target_columns = df.columns.tolist()
    
    result = {}
    for col in target_columns:
        # 转为字符串并去除 NaN，获取唯一值
        unique_values = df[col].dropna().astype(str).unique()
        result[col] = set(unique_values)
    
    return result


def main():
    parser = argparse.ArgumentParser(description="分析 Excel 文件每列的唯一值")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--header-row", type=int, default=0, help="表头所在行，默认 0")
    parser.add_argument("-s", "--sheet", default=0, help="工作表名称或索引，默认 0")
    parser.add_argument("-c", "--columns", nargs="+", help="指定要分析的列名（可多个）")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    try:
        result = analyze_excel_columns(
            args.file,
            header_row=args.header_row,
            columns=args.columns,
            sheet_name=sheet,
        )
        
        if args.json:
            # 转换 set 为 list 以便 JSON 序列化
            json_result = {k: sorted(v) for k, v in result.items()}
            print(json.dumps(json_result, ensure_ascii=False, indent=2))
        else:
            for col, values in result.items():
                print(f"\n【{col}】({len(values)} 个唯一值)")
                print(f"  {sorted(values)}")
                
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
