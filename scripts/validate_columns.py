"""
校验 Excel 列名是否符合预期模板
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# 考勤表标准列名模板
ATTENDANCE_COLUMNS = [
    "工号", "部门", "人员类型", "员工状态", "入职日期", "离职日期",
    "日期", "星期", "班次", "考勤组",
    "上班 1 打卡时间", "上班 1 打卡结果", "上班 1 修改原因", "上班 1 打卡地点",
    "下班 1 打卡时间", "下班 1 打卡结果", "下班 1 修改原因", "下班 1 打卡地点",
    "应出勤天数", "应出勤时长(小时)", "休息或未排班天数",
    "实际出勤天数", "实际出勤时长(小时)", "班内工作时长(小时)",
    "出差天数", "外出时长", "补卡次数",
    "迟到次数", "迟到时长(小时)", "严重迟到次数", "严重迟到时长(小时)",
    "早退次数", "早退时长(小时)", "缺勤时长(小时)",
    "上班缺卡次数", "下班缺卡次数", "旷工天数",
    "加班总时长(小时)", "加班总时长 - 计加班费(小时)", "加班总时长 - 计调休(小时)",
]


def validate_columns(
    file_path: str,
    header_row: int = 0,
    required_columns: list[str] | None = None,
    sheet_name: str | int = 0,
) -> dict:
    """
    校验 Excel 列名是否符合预期模板
    
    Args:
        file_path: Excel 文件路径
        header_row: 表头所在行（从 0 开始）
        required_columns: 必需的列名列表，为 None 时使用考勤表模板
        sheet_name: 工作表名称或索引，默认第一个 sheet
    
    Returns:
        校验结果字典，包含 valid, missing, extra, matched
    """
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"文件不存在: {file_path}")
    
    if required_columns is None:
        required_columns = ATTENDANCE_COLUMNS
    
    df = pd.read_excel(file_path, header=header_row, nrows=0, sheet_name=sheet_name)
    actual_columns = [str(c).strip() for c in df.columns.tolist()]
    
    required_set = set(required_columns)
    actual_set = set(actual_columns)
    
    matched = required_set & actual_set
    missing = required_set - actual_set
    extra = actual_set - required_set
    
    return {
        "valid": len(missing) == 0,
        "matched": sorted(matched),
        "missing": sorted(missing),
        "extra": sorted(extra),
        "match_rate": len(matched) / len(required_set) if required_set else 1.0,
    }


def main():
    parser = argparse.ArgumentParser(description="校验 Excel 列名是否符合模板")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--header-row", type=int, default=0, help="表头所在行，默认 0")
    parser.add_argument("-s", "--sheet", default="0", help="工作表名称或索引，默认 0")
    
    args = parser.parse_args()
    sheet = int(args.sheet) if args.sheet.isdigit() else args.sheet
    
    try:
        result = validate_columns(args.file, header_row=args.header_row, sheet_name=sheet)
        
        print(f"匹配率: {result['match_rate']:.1%}")
        print(f"匹配列数: {len(result['matched'])}")
        
        if result["missing"]:
            print(f"\n缺失列 ({len(result['missing'])} 个):")
            for col in result["missing"]:
                print(f"  - {col}")
        
        if result["extra"]:
            print(f"\n额外列 ({len(result['extra'])} 个):")
            for col in result["extra"]:
                print(f"  + {col}")
        
        if result["valid"]:
            print("\n✓ 校验通过")
        else:
            print("\n✗ 校验失败：存在缺失列")
            sys.exit(1)
            
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
