# import pytest
from src.userSetting import UserSetting
from dataclasses import asdict

test_setting = dict(
    studentID="x00x000x",
    userName="testName",
    faculty="testFaculty",
    roomNumber="testNumber",
    excelDir="C:\\Users\\color\\Desktop",
)


def test_excelFileName():
    us = UserSetting(**test_setting)
    assert us.excelFileName(3, 5) == "研究活動状況表(R3.5)_testName.xlsx"


def test_asdict():
    us = UserSetting(**test_setting)
    assert asdict(us) == test_setting
