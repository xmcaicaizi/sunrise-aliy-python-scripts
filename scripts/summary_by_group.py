"""
按指定维度分组汇总考勤统计
支持：个体（工号）、部门、地区、地区-部门等多维度
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from detect_header import detect_header_row

# 默认汇总字段配置
DEFAULT_SUM_COLUMNS = [
    "实际出勤天数",
    "迟到次数",
    "严重迟到次数",
    "早退次数",
    "上班缺卡次数",
    "下班缺卡次数",
    "旷工天数",
    "补卡次数",
]


def summary_by_group(
    file_path: str,
    group_by: list[str],
    header_row: int | None = None,
    sum_columns: list[str] | None = None,
    output_path: str | None = None,
    auto_detect_header: bool = True,
    sheet_name: str | int = 0,
) -> pd.DataFrame:
    """
    按指定维度分组汇总考勤统计
    
    Args:
        file_path: Excel 文件路径
        group_by: 分组列名列表（如 ["部门"] 或 ["地区", "部门"]）
        header_row: 表头所在行，为 None 时自动检测
        sum_columns: 要汇总的列名列表，为 None 时使用默认配置
        output_path: 输出文件路径，为 None 时不保存
        auto_detect_header: 是否自动检测表头行
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        汇总后的 DataFrame
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
    
    # 检查分组列是否存在
    missing_cols = [c for c in group_by if c not in df.columns]
    if missing_cols:
        raise ValueError(f"分组列不存在: {missing_cols}。可用列名: {list(df.columns)}")
    
    if sum_columns is None:
        sum_columns = DEFAULT_SUM_COLUMNS
    
    # 过滤出存在的汇总列
    existing_sum_cols = [c for c in sum_columns if c in df.columns]
    missing_sum_cols = [c for c in sum_columns if c not in df.columns]
    if missing_sum_cols:
        print(f"警告: 以下汇总列不存在，已跳过: {missing_sum_cols}")
    
    if not existing_sum_cols:
        raise ValueError("没有可用的汇总列")
    
    # 构建聚合规则
    agg_dict = {col: "sum" for col in existing_sum_cols}
    
    # 添加人数统计（如果有工号列）
    if "工号" in df.columns and "工号" not in group_by:
        agg_dict["工号"] = "nunique"
    
    # 按分组列汇总
    result = df.groupby(group_by, as_index=False).agg(agg_dict)
    
    # 重命名工号列为人数
    if "工号" in result.columns and "工号" not in group_by:
        result = result.rename(columns={"工号": "人数"})
    
    # 计算人均指标
    if "人数" in result.columns:
        for col in existing_sum_cols:
            result[f"人均{col}"] = (result[col] / result["人数"]).round(2)
    
    print(f"分组维度: {group_by}")
    print(f"汇总字段: {existing_sum_cols}")
    print(f"共 {len(result)} 条记录")
    
    if output_path:
        result.to_excel(output_path, index=False)
        print(f"已保存到: {output_path}")
    
    return result


def main():
    parser = argparse.ArgumentParser(description="按指定维度分组汇总考勤统计")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument(
        "-g", "--group-by",
        nargs="+",
        required=True,
        help="分组列名（可多个，如 -g 部门 或 -g 地区 部门）",
    )
    parser.add_argument("--header-row", type=int, help="表头所在行（不指定则自动检测）")
    parser.add_argument("-s", "--sheet", default="0", help="工作表名称或索引，默认 0")
    parser.add_argument("-c", "--columns", nargs="+", help="要汇总的列名（不指定则使用默认配置）")
    parser.add_argument("-o", "--output", help="输出文件路径")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    try:
        summary_by_group(
            args.file,
            group_by=args.group_by,
            header_row=args.header_row,
            sum_columns=args.columns,
            output_path=args.output,
            sheet_name=sheet,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
