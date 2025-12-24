"""
读取 Excel 文件前五行，用于判断表头结构（是否为多级表头）
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


def read_excel_head(file_path: str, rows: int = 5) -> pd.DataFrame:
    """
    读取 Excel 文件的前 N 行
    
    Args:
        file_path: Excel 文件路径
        rows: 读取的行数，默认 5 行
    
    Returns:
        包含前 N 行数据的 DataFrame
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    if not path.suffix.lower() in ['.xlsx', '.xls']:
        raise ValueError(f"不支持的文件格式: {path.suffix}")
    
    # 读取时不指定 header，保留原始数据结构
    df = pd.read_excel(file_path, header=None, nrows=rows)
    return df


def main():
    parser = argparse.ArgumentParser(description="读取 Excel 文件前五行")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("-n", "--rows", type=int, default=5, help="读取行数，默认 5")
    
    args = parser.parse_args()
    
    try:
        df = read_excel_head(args.file, args.rows)
        print(f"文件: {args.file}")
        print(f"前 {args.rows} 行数据:")
        print("-" * 50)
        print(df.to_string())
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
