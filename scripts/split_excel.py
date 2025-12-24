"""
按指定列拆分 Excel 文件，每个唯一值生成一个单独的文件
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from detect_header import detect_header_row


def split_excel(
    file_path: str,
    column: str,
    header_row: int | None = None,
    output_dir: str | None = None,
    auto_detect_header: bool = True,
    sheet_name: str | int = 0,
) -> dict[str, int]:
    """
    按指定列拆分 Excel 文件
    
    Args:
        file_path: Excel 文件路径
        column: 用于拆分的列名
        header_row: 表头所在行，为 None 时自动检测
        output_dir: 输出目录，为 None 时使用源文件所在目录
        auto_detect_header: 是否自动检测表头行
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        字典，key 为拆分值，value 为该文件的行数
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    # 自动检测表头行
    if header_row is None and auto_detect_header:
        header_row = detect_header_row(file_path, sheet_name=sheet_name)
        print(f"自动检测表头行: {header_row}")
    elif header_row is None:
        header_row = 0
    
    df = pd.read_excel(file_path, header=header_row, sheet_name=sheet_name)
    
    if column not in df.columns:
        raise ValueError(f"列名 '{column}' 不存在。可用列名: {list(df.columns)}")
    
    # 确定输出目录
    if output_dir is None:
        out_path = path.parent / f"{path.stem}_split"
    else:
        out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    
    # 按列值分组并导出
    result = {}
    unique_values = df[column].dropna().unique()
    
    for value in unique_values:
        subset = df[df[column] == value]
        # 清理文件名中的非法字符
        safe_name = str(value).replace("/", "_").replace("\\", "_").replace(":", "_")
        output_file = out_path / f"{safe_name}.xlsx"
        subset.to_excel(output_file, index=False)
        result[str(value)] = len(subset)
        print(f"导出 [{value}]: {len(subset)} 行 -> {output_file}")
    
    print(f"\n共拆分为 {len(result)} 个文件，保存在: {out_path}")
    return result


def main():
    parser = argparse.ArgumentParser(description="按指定列拆分 Excel 文件")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("-c", "--column", required=True, help="用于拆分的列名")
    parser.add_argument("--header-row", type=int, help="表头所在行（不指定则自动检测）")
    parser.add_argument("-s", "--sheet", default="0", help="工作表名称或索引，默认 0")
    parser.add_argument("-o", "--output-dir", help="输出目录")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    try:
        split_excel(
            args.file,
            args.column,
            header_row=args.header_row,
            output_dir=args.output_dir,
            sheet_name=sheet,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
