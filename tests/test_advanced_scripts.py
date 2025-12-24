"""
高级脚本功能测试
"""

import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))

from clean_attendance import clean_attendance
from summary_by_employee import summary_by_employee
from summary_by_group import summary_by_group

TEST_FILE = Path(__file__).parent.parent / "examples" / "test01.xlsx"


@pytest.fixture
def test_file():
    if not TEST_FILE.exists():
        pytest.skip(f"测试文件不存在: {TEST_FILE}")
    return str(TEST_FILE)


class TestCleanAttendance:
    """clean_attendance.py 测试"""

    def test_clean_with_default_rules(self, test_file):
        """测试默认清洗规则"""
        df = clean_attendance(test_file)
        # 清洗后应该没有周末
        assert "星期六" not in df["星期"].values
        assert "星期日" not in df["星期"].values
        # 清洗后应该没有离职员工
        assert "离职" not in df["员工状态"].values

    def test_clean_preserves_valid_data(self, test_file):
        """测试清洗后保留有效数据"""
        df = clean_attendance(test_file)
        # 应该有数据剩余
        assert len(df) > 0
        # 应该保留正式员工
        assert "正式" in df["人员类型"].values


class TestSummaryByEmployee:
    """summary_by_employee.py 测试"""

    def test_summary_default_columns(self, test_file):
        """测试默认汇总字段"""
        df = summary_by_employee(test_file)
        assert "工号" in df.columns
        assert "部门" in df.columns
        assert "实际出勤天数" in df.columns
        assert "迟到次数" in df.columns

    def test_summary_has_unique_employees(self, test_file):
        """测试汇总后工号唯一"""
        df = summary_by_employee(test_file)
        assert df["工号"].is_unique


class TestSummaryByGroup:
    """summary_by_group.py 测试"""

    def test_summary_by_department(self, test_file):
        """测试按部门汇总"""
        df = summary_by_group(test_file, group_by=["部门"])
        assert "部门" in df.columns
        assert "人数" in df.columns
        assert "实际出勤天数" in df.columns
        # 应该有人均指标
        assert "人均实际出勤天数" in df.columns

    def test_summary_multi_dimension(self, test_file):
        """测试多维度汇总"""
        df = summary_by_group(test_file, group_by=["部门", "人员类型"])
        assert "部门" in df.columns
        assert "人员类型" in df.columns
