"""
脚本功能测试
使用 examples/test01.xlsx 作为测试数据
"""

import sys
from pathlib import Path

import pandas as pd
import pytest

# 添加 scripts 目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from analyze_excel_columns import analyze_excel_columns
from detect_header import detect_header_row
from filter_excel import filter_excel
from read_excel_head import read_excel_head
from validate_columns import validate_columns

# 测试数据路径
TEST_FILE = Path(__file__).parent.parent / "examples" / "test01.xlsx"


@pytest.fixture
def test_file():
    """确保测试文件存在"""
    if not TEST_FILE.exists():
        pytest.skip(f"测试文件不存在: {TEST_FILE}")
    return str(TEST_FILE)


class TestReadExcelHead:
    """read_excel_head.py 测试"""

    def test_read_default_rows(self, test_file):
        """测试默认读取 5 行"""
        result = read_excel_head(test_file)
        if isinstance(result, dict):
            # 多 sheet 情况
            for df in result.values():
                assert len(df) <= 5
        else:
            assert len(result) <= 5

    def test_read_custom_rows(self, test_file):
        """测试自定义行数"""
        result = read_excel_head(test_file, rows=3)
        if isinstance(result, dict):
            for df in result.values():
                assert len(df) <= 3
        else:
            assert len(result) <= 3


class TestDetectHeader:
    """detect_header.py 测试"""

    def test_detect_header_row(self, test_file):
        """测试表头检测"""
        header_row = detect_header_row(test_file)
        # test01.xlsx 的真实表头在第 2 行（索引 1）
        assert header_row == 1

    def test_detect_with_custom_keywords(self, test_file):
        """测试自定义关键字"""
        header_row = detect_header_row(test_file, keywords=["工号", "部门"])
        assert header_row == 1


class TestValidateColumns:
    """validate_columns.py 测试"""

    def test_validate_attendance_columns(self, test_file):
        """测试考勤表列名校验"""
        result = validate_columns(test_file, header_row=1)
        assert result["valid"] is True
        assert result["match_rate"] == 1.0
        assert len(result["missing"]) == 0

    def test_validate_with_custom_columns(self, test_file):
        """测试自定义列名校验"""
        result = validate_columns(
            test_file,
            header_row=1,
            required_columns=["工号", "部门", "不存在的列"],
        )
        assert "不存在的列" in result["missing"]
        assert "工号" in result["matched"]
        assert "部门" in result["matched"]


class TestAnalyzeExcelColumns:
    """analyze_excel_columns.py 测试"""

    def test_analyze_all_columns(self, test_file):
        """测试分析所有列"""
        result = analyze_excel_columns(test_file, header_row=1)
        assert isinstance(result, dict)
        assert "工号" in result
        assert "部门" in result

    def test_analyze_specific_columns(self, test_file):
        """测试分析指定列"""
        result = analyze_excel_columns(
            test_file,
            header_row=1,
            columns=["星期", "人员类型"],
        )
        assert len(result) == 2
        assert "星期" in result
        assert "人员类型" in result
        # 星期应该有 7 个唯一值
        assert len(result["星期"]) == 7


class TestFilterExcel:
    """filter_excel.py 测试"""

    def test_filter_weekend(self, test_file):
        """测试剔除周末"""
        df = filter_excel(
            test_file,
            column="星期",
            values=["星期六", "星期日"],
            header_row=1,
        )
        # 确保没有周末数据
        assert "星期六" not in df["星期"].values
        assert "星期日" not in df["星期"].values

    def test_filter_employee_type(self, test_file):
        """测试剔除人员类型"""
        df = filter_excel(
            test_file,
            column="人员类型",
            values=["实习", "外包"],
            header_row=1,
        )
        assert "实习" not in df["人员类型"].values
        assert "外包" not in df["人员类型"].values
