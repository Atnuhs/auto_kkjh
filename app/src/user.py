# from datetime import date
from pathlib import Path
from dataclasses import dataclass


@dataclass(frozen=True)
class StudentId:
    value: str


@dataclass(frozen=True)
class UserName:
    value: str


@dataclass(frozen=True)
class Faculty:
    value: str


@dataclass(frozen=True)
class RoomNumber:
    value: str


@dataclass(frozen=True)
class ExcelDir:
    value: Path


@dataclass
class User:
    studentId: StudentId
    userName: UserName
    faculty: Faculty
    roomNumber: RoomNumber
    excelDir: ExcelDir

    def __eq__(self, other):
        return isinstance(other, User) and self.studentId == other.studentId

    def changeName(self, newUserName: UserName):
        self.uerName = newUserName

    def changeFaculty(self, newFaculty: Faculty):
        self.faculty = newFaculty

    def changeRoomNumber(self, newRoomNumber: RoomNumber):
        self.roomNumber = newRoomNumber

    def changeExcelDir(self, newExcelDir: ExcelDir):
        self.excelDir = newExcelDir

    def excelFileName(self, reiwa_year: int, month: int):
        fileSuffix = f"(R{reiwa_year}.{month})_{self.userName}.xlsx"
        return f"研究活動状況表{fileSuffix}"

    def excelPath(self, reiwa_year: int, month: int):
        excelDir = self.excelDir
        return excelDir / self.excelFileName(reiwa_year, month)
