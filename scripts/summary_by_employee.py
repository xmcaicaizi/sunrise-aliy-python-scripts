"""
按工号汇总考勤统计
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

# 需要保留的基础信息列（取第一条记录的值）
INFO_COLUMNS = ["部门", "人员类型", "员工状态"]


def summary_by_employee(
    file_path: str,
    header_row: int | None = None,
    sum_columns: list[str] | None = None,
    output_path: str | None = None,
    auto_detect_header: bool = True,
) -> pd.DataFrame:
    """
    按工号汇总考勤统计
    
    Args:
        file_path: Excel 文件路径
        header_row: 表头所在行，为 None 时自动检测
        sum_columns: 要汇总的列名列表，为 None 时使用默认配置
        output_path: 输出文件路径，为 None 时不保存
        auto_detect_header: 是否自动检测表头行
    
    Returns:
        汇总后的 DataFrame
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    # 自动检测表头行
    if header_row is None and auto_detect_header:
        header_row = detect_header_row(file_path)
        print(f"自动检测表头行: {header_row}")
    elif header_row is None:
        header_row = 0
    
    df = pd.read_excel(file_path, header=header_row)
    
    if "工号" not in df.columns:
        raise ValueError("数据中缺少'工号'列")
    
    if sum_columns is None:
        sum_columns = DEFAULT_SUM_COLUMNS
    
    # 过滤出存在的汇总列
    existing_sum_cols = [c for c in sum_columns if c in df.columns]
    missing_cols = [c for c in sum_columns if c not in df.columns]
    if missing_cols:
        print(f"警告: 以下列不存在，已跳过: {missing_cols}")
    
    # 过滤出存在的信息列
    existing_info_cols = [c for c in INFO_COLUMNS if c in df.columns]
    
    # 构建聚合规则
    agg_dict = {}
    for col in existing_sum_cols:
        agg_dict[col] = "sum"
    for col in existing_info_cols:
        agg_dict[col] = "first"
    
    # 按工号分组汇总
    result = df.groupby("工号", as_index=False).agg(agg_dict)
    
    # 调整列顺序：工号 + 信息列 + 汇总列
    col_order = ["工号"] + existing_info_cols + existing_sum_cols
    result = result[col_order]
    
    print(f"共汇总 {len(result)} 名员工")
    print(f"汇总字段: {existing_sum_cols}")
    
    if output_path:
        result.to_excel(output_path, index=False)
        print(f"已保存到: {output_path}")
    
    return result


def main():
    parser = argparse.ArgumentParser(description="按工号汇总考勤统计")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--header-row", type=int, help="表头所在行（不指定则自动检测）")
    parser.add_argument("-c", "--columns", nargs="+", help="要汇总的列名（不指定则使用默认配置）")
    parser.add_argument("-o", "--output", help="输出文件路径")
    
    args = parser.parse_args()
    
    try:
        summary_by_employee(
            args.file,
            header_row=args.header_row,
            sum_columns=args.columns,
            output_path=args.output,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
