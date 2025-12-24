"""
通过指定列关联两个 Excel 文件
典型场景：将花名册中的地区信息关联到考勤数据
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from detect_header import detect_header_row


def join_excel(
    left_file: str,
    right_file: str,
    on: str,
    right_columns: list[str] | None = None,
    left_header_row: int | None = None,
    right_header_row: int | None = None,
    left_sheet: str | int = 0,
    right_sheet: str | int = 0,
    output_path: str | None = None,
    how: str = "left",
) -> pd.DataFrame:
    """
    通过指定列关联两个 Excel 文件
    
    Args:
        left_file: 左表文件路径（主表，如考勤数据）
        right_file: 右表文件路径（关联表，如花名册）
        on: 关联列名（两表中必须都存在）
        right_columns: 从右表中选取的列名列表，为 None 时选取所有列
        left_header_row: 左表表头行，为 None 时自动检测
        right_header_row: 右表表头行，为 None 时自动检测
        left_sheet: 左表工作表
        right_sheet: 右表工作表
        output_path: 输出文件路径
        how: 关联方式，默认 left（保留左表所有行）
    
    Returns:
        关联后的 DataFrame
    """
    # 检查文件存在
    if not Path(left_file).exists():
        raise FileNotFoundError(f"左表文件不存在: {left_file}")
    if not Path(right_file).exists():
        raise FileNotFoundError(f"右表文件不存在: {right_file}")
    
    # 自动检测表头行
    if left_header_row is None:
        left_header_row = detect_header_row(left_file, sheet_name=left_sheet)
        print(f"左表自动检测表头行: {left_header_row}")
    
    if right_header_row is None:
        right_header_row = detect_header_row(right_file, sheet_name=right_sheet)
        print(f"右表自动检测表头行: {right_header_row}")
    
    # 读取数据
    df_left = pd.read_excel(left_file, header=left_header_row, sheet_name=left_sheet)
    df_right = pd.read_excel(right_file, header=right_header_row, sheet_name=right_sheet)
    
    # 检查关联列
    if on not in df_left.columns:
        raise ValueError(f"左表中不存在关联列 '{on}'。可用列: {list(df_left.columns)}")
    if on not in df_right.columns:
        raise ValueError(f"右表中不存在关联列 '{on}'。可用列: {list(df_right.columns)}")
    
    # 统一关联列类型为字符串，并补齐前导零（工号场景）
    df_left[on] = df_left[on].astype(str).str.strip()
    df_right[on] = df_right[on].astype(str).str.strip()
    
    # 如果是工号，尝试统一格式（补齐前导零到6位）
    if on == "工号":
        df_left[on] = df_left[on].str.zfill(6)
        df_right[on] = df_right[on].str.zfill(6)
    
    # 选取右表列
    if right_columns:
        missing = [c for c in right_columns if c not in df_right.columns]
        if missing:
            raise ValueError(f"右表中不存在列: {missing}。可用列: {list(df_right.columns)}")
        # 确保包含关联列
        select_cols = [on] + [c for c in right_columns if c != on]
        df_right = df_right[select_cols].drop_duplicates(subset=[on])
    else:
        df_right = df_right.drop_duplicates(subset=[on])
    
    print(f"左表: {len(df_left)} 行, {len(df_left.columns)} 列")
    print(f"右表: {len(df_right)} 行, {len(df_right.columns)} 列")
    
    # 关联
    result = pd.merge(df_left, df_right, on=on, how=how, suffixes=("", "_右表"))
    
    print(f"关联后: {len(result)} 行, {len(result.columns)} 列")
    
    # 统计关联情况
    if how == "left":
        # 检查有多少行没有匹配到
        new_cols = [c for c in result.columns if c not in df_left.columns]
        if new_cols:
            null_count = result[new_cols[0]].isna().sum()
            print(f"未匹配行数: {null_count}")
    
    if output_path:
        result.to_excel(output_path, index=False)
        print(f"已保存到: {output_path}")
    
    return result


def main():
    parser = argparse.ArgumentParser(description="通过指定列关联两个 Excel 文件")
    parser.add_argument("left_file", help="左表文件路径（主表）")
    parser.add_argument("right_file", help="右表文件路径（关联表）")
    parser.add_argument("--on", required=True, help="关联列名")
    parser.add_argument("-c", "--columns", nargs="+", help="从右表选取的列名（不指定则选取所有）")
    parser.add_argument("--left-header-row", type=int, help="左表表头行")
    parser.add_argument("--right-header-row", type=int, help="右表表头行")
    parser.add_argument("--left-sheet", default="0", help="左表工作表")
    parser.add_argument("--right-sheet", default="0", help="右表工作表")
    parser.add_argument("--how", default="left", choices=["left", "inner", "outer"], help="关联方式")
    parser.add_argument("-o", "--output", help="输出文件路径")
    
    args = parser.parse_args()
    
    left_sheet = int(args.left_sheet) if args.left_sheet.isdigit() else args.left_sheet
    right_sheet = int(args.right_sheet) if args.right_sheet.isdigit() else args.right_sheet
    
    try:
        join_excel(
            args.left_file,
            args.right_file,
            on=args.on,
            right_columns=args.columns,
            left_header_row=args.left_header_row,
            right_header_row=args.right_header_row,
            left_sheet=left_sheet,
            right_sheet=right_sheet,
            output_path=args.output,
            how=args.how,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
