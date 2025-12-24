"""
考勤数据清洗一站式脚本
整合剔除周末、非正式员工、离职员工等默认规则
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from detect_header import detect_header_row

# 默认清洗规则
# 无需打卡类型：休息、出差、自由班制、请假、补卡通过
DEFAULT_RULES = {
    "星期": ["星期六", "星期日"],  # 剔除周末
    "人员类型": ["实习", "外包"],  # 剔除非正式员工
    "员工状态": ["离职"],  # 剔除离职员工
    "上班 1 打卡结果": [
        "缺卡",
        "无需打卡(休息)",
        "无需打卡(出差)",
        "无需打卡(自由班制)",
        "无需打卡(请假)",
    ],
    "下班 1 打卡结果": [
        "缺卡",
        "无需打卡(休息)",
        "无需打卡(出差)",
        "无需打卡(自由班制)",
        "无需打卡(请假)",
        "无需打卡(补卡通过)",
    ],
}


def clean_attendance(
    file_path: str,
    header_row: int | None = None,
    rules: dict[str, list[str]] | None = None,
    output_path: str | None = None,
    auto_detect_header: bool = True,
    sheet_name: str | int = 0,
) -> pd.DataFrame:
    """
    考勤数据清洗
    
    Args:
        file_path: Excel 文件路径
        header_row: 表头所在行，为 None 时自动检测
        rules: 清洗规则字典，key 为列名，value 为要剔除的值列表
        output_path: 输出文件路径，为 None 时不保存
        auto_detect_header: 是否自动检测表头行
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        清洗后的 DataFrame
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
    original_count = len(df)
    
    if rules is None:
        rules = DEFAULT_RULES
    
    # 应用清洗规则
    stats = {}
    for column, values in rules.items():
        if column not in df.columns:
            print(f"警告: 列 '{column}' 不存在，跳过该规则")
            continue
        
        before = len(df)
        mask = ~df[column].astype(str).isin(values)
        df = df[mask].copy()
        removed = before - len(df)
        stats[column] = removed
        
        if removed > 0:
            print(f"剔除 [{column}] 包含 {values}: {removed} 行")
    
    print(f"\n清洗统计:")
    print(f"  原始行数: {original_count}")
    print(f"  剔除行数: {original_count - len(df)}")
    print(f"  剩余行数: {len(df)}")
    
    if output_path:
        df.to_excel(output_path, index=False)
        print(f"\n已保存到: {output_path}")
    
    return df


def main():
    parser = argparse.ArgumentParser(description="考勤数据清洗一站式脚本")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--header-row", type=int, help="表头所在行（不指定则自动检测）")
    parser.add_argument("-s", "--sheet", default="0", help="工作表名称或索引，默认 0")
    parser.add_argument("-o", "--output", help="输出文件路径")
    parser.add_argument("--no-weekend", action="store_true", help="不剔除周末")
    parser.add_argument("--no-intern", action="store_true", help="不剔除实习/外包")
    parser.add_argument("--no-resigned", action="store_true", help="不剔除离职员工")
    parser.add_argument("--no-abnormal", action="store_true", help="不剔除异常打卡")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    # 根据参数调整规则
    rules = DEFAULT_RULES.copy()
    if args.no_weekend:
        rules.pop("星期", None)
    if args.no_intern:
        rules.pop("人员类型", None)
    if args.no_resigned:
        rules.pop("员工状态", None)
    if args.no_abnormal:
        rules.pop("上班 1 打卡结果", None)
        rules.pop("下班 1 打卡结果", None)
    
    try:
        clean_attendance(
            args.file,
            header_row=args.header_row,
            rules=rules,
            output_path=args.output,
            sheet_name=sheet,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
