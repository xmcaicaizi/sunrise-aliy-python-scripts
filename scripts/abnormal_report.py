"""
异常考勤报告生成
筛选缺卡、旷工、严重迟到等异常记录
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

from detect_header import detect_header_row

# 默认异常条件
DEFAULT_ABNORMAL_CONDITIONS = {
    "缺卡": {
        "columns": ["上班 1 打卡结果", "下班 1 打卡结果"],
        "values": ["缺卡"],
    },
    "旷工": {
        "columns": ["旷工天数"],
        "condition": "gt",  # greater than
        "threshold": 0,
    },
    "严重迟到": {
        "columns": ["严重迟到次数"],
        "condition": "gt",
        "threshold": 0,
    },
    "迟到": {
        "columns": ["迟到次数"],
        "condition": "gt",
        "threshold": 0,
    },
    "早退": {
        "columns": ["早退次数"],
        "condition": "gt",
        "threshold": 0,
    },
}


def filter_abnormal(
    df: pd.DataFrame,
    abnormal_type: str,
    config: dict,
) -> pd.DataFrame:
    """根据配置筛选异常记录"""
    columns = config["columns"]
    
    # 检查列是否存在
    existing_cols = [c for c in columns if c in df.columns]
    if not existing_cols:
        return pd.DataFrame()
    
    if "values" in config:
        # 值匹配模式
        mask = pd.Series([False] * len(df))
        for col in existing_cols:
            mask |= df[col].astype(str).isin(config["values"])
        return df[mask].copy()
    
    elif "condition" in config:
        # 数值比较模式
        mask = pd.Series([False] * len(df))
        threshold = config["threshold"]
        for col in existing_cols:
            if config["condition"] == "gt":
                mask |= pd.to_numeric(df[col], errors="coerce") > threshold
            elif config["condition"] == "gte":
                mask |= pd.to_numeric(df[col], errors="coerce") >= threshold
        return df[mask].copy()
    
    return pd.DataFrame()


def generate_abnormal_report(
    file_path: str,
    header_row: int | None = None,
    abnormal_types: list[str] | None = None,
    output_path: str | None = None,
    auto_detect_header: bool = True,
) -> dict[str, pd.DataFrame]:
    """
    生成异常考勤报告
    
    Args:
        file_path: Excel 文件路径
        header_row: 表头所在行，为 None 时自动检测
        abnormal_types: 要筛选的异常类型列表，为 None 时筛选所有类型
        output_path: 输出文件路径，为 None 时不保存
        auto_detect_header: 是否自动检测表头行
    
    Returns:
        字典，key 为异常类型，value 为对应的 DataFrame
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
    
    if abnormal_types is None:
        abnormal_types = list(DEFAULT_ABNORMAL_CONDITIONS.keys())
    
    results = {}
    all_abnormal = pd.DataFrame()
    
    for abnormal_type in abnormal_types:
        if abnormal_type not in DEFAULT_ABNORMAL_CONDITIONS:
            print(f"警告: 未知的异常类型 '{abnormal_type}'，跳过")
            continue
        
        config = DEFAULT_ABNORMAL_CONDITIONS[abnormal_type]
        abnormal_df = filter_abnormal(df, abnormal_type, config)
        
        if not abnormal_df.empty:
            abnormal_df = abnormal_df.copy()
            abnormal_df["异常类型"] = abnormal_type
            results[abnormal_type] = abnormal_df
            all_abnormal = pd.concat([all_abnormal, abnormal_df], ignore_index=True)
            print(f"【{abnormal_type}】: {len(abnormal_df)} 条记录")
        else:
            print(f"【{abnormal_type}】: 0 条记录")
    
    print(f"\n异常记录总数: {len(all_abnormal)}")
    
    if output_path and not all_abnormal.empty:
        all_abnormal.to_excel(output_path, index=False)
        print(f"已保存到: {output_path}")
    
    return results


def main():
    parser = argparse.ArgumentParser(description="生成异常考勤报告")
    parser.add_argument("file", help="Excel 文件路径")
    parser.add_argument("--header-row", type=int, help="表头所在行（不指定则自动检测）")
    parser.add_argument(
        "-t", "--types",
        nargs="+",
        choices=list(DEFAULT_ABNORMAL_CONDITIONS.keys()),
        help="要筛选的异常类型",
    )
    parser.add_argument("-o", "--output", help="输出文件路径")
    
    args = parser.parse_args()
    
    try:
        generate_abnormal_report(
            args.file,
            header_row=args.header_row,
            abnormal_types=args.types,
            output_path=args.output,
        )
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
