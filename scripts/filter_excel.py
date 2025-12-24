"""
根据指定的列名和值剔除 Excel 数据行
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


def filter_excel(
    file_path: str,
    column: str,
    values: list[str],
    header_row: int = 0,
    output_path: str | None = None,
    sheet_name: str | int = 0,
) -> pd.DataFrame:
    """
    剔除 Excel 中指定列包含特定值的行
    
    Args:
        file_path: Excel 文件路径
        column: 列名
        values: 要剔除的值列表
        header_row: 表头所在行（从 0 开始），默认第 0 行
        output_path: 输出文件路径，为 None 时不保存
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        过滤后的 DataFrame
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    df = pd.read_excel(file_path, header=header_row, sheet_name=sheet_name)
    
    if column not in df.columns:
        raise ValueError(f"列名 '{column}' 不存在。可用列名: {list(df.columns)}")
    
    original_count = len(df)
    
    # 剔除包含指定值的行
    mask = ~df[column].astype(str).isin(values)
    df_filtered = df[mask].copy()
    
    removed_count = original_count - len(df_filtered)
    print(f"原始行数: {original_count}")
    print(f"剔除行数: {removed_count}")
    print(f"剩余行数: {len(df_filtered)}")
    
    if output_path:
        df_filtered.to_excel(output_path, index=False)
        print(f"已保存到: {output_path}")
    
    return df_filtered


def main():
    parser = argparse.ArgumentParser(description="剔除 Excel 中指定列包含特定值的行")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("-c", "--column", required=True, help="列名")
    parser.add_argument("-v", "--values", nargs="+", required=True, help="要剔除的值（可多个）")
    parser.add_argument("--header-row", type=int, default=0, help="表头所在行，默认 0")
    parser.add_argument("-s", "--sheet", default="0", help="工作表名称或索引，默认 0")
    parser.add_argument("-o", "--output", help="输出文件路径")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    try:
        filter_excel(
            args.file,
            args.column,
            args.values,
            header_row=args.header_row,
            output_path=args.output,
            sheet_name=sheet,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
