"""
读取 Excel 文件前五行，用于判断表头结构（是否为多级表头）
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


def read_excel_head(
    file_path: str,
    rows: int = 5,
    sheet_name: str | int | None = None,
) -> pd.DataFrame | dict[str, pd.DataFrame]:
    """
    读取 Excel 文件的前 N 行
    
    Args:
        file_path: Excel 文件路径
        rows: 读取的行数，默认 5 行
        sheet_name: 工作表名称或索引，为 None 时读取所有非空 sheet
    
    Returns:
        单个 sheet 返回 DataFrame，多个 sheet 返回字典
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    if path.suffix.lower() not in ['.xlsx', '.xls']:
        raise ValueError(f"不支持的文件格式: {path.suffix}")
    
    if sheet_name is not None:
        # 读取指定 sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=rows)
        return df
    
    # 读取所有 sheet
    xlsx = pd.ExcelFile(file_path)
    results = {}
    for name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=name, header=None, nrows=rows)
        if not df.empty:
            results[name] = df
    
    if len(results) == 1:
        return list(results.values())[0]
    return results


def main():
    parser = argparse.ArgumentParser(description="读取 Excel 文件前五行")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("-n", "--rows", type=int, default=5, help="读取行数，默认 5")
    parser.add_argument("-s", "--sheet", help="工作表名称（不指定则读取所有非空 sheet）")
    
    args = parser.parse_args()
    
    try:
        result = read_excel_head(args.file, args.rows, args.sheet)
        print(f"文件: {args.file}")
        
        if isinstance(result, dict):
            print(f"共 {len(result)} 个非空工作表")
            for name, df in result.items():
                print(f"\n【{name}】前 {args.rows} 行:")
                print("-" * 50)
                print(df.to_string())
        else:
            print(f"前 {args.rows} 行数据:")
            print("-" * 50)
            if result.empty:
                # 尝试列出所有 sheet 名称
                xlsx = pd.ExcelFile(args.file)
                print(f"数据为空。可用工作表: {xlsx.sheet_names}")
            else:
                print(result.to_string())
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
