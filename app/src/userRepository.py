import json
import sqlite3
from typing import Union
from src.user import User, UserName, StudentId, Faculty, RoomNumber, ExcelDir


class UserRepository:
    def find(self, userName: UserName) -> Union[User, None]:
        conn = sqlite3.connect("sample.db")
        conn.row_factory = sqlite3.Row
        c = conn.cursor()
        t = (userName.value,)
        c.execute("SELECT * FROM users WHERE uerName=?", t)
        userDict = c.fetchone()
        if userDict:
            studentId = StudentId(userDict["studentId"])
            faculty = Faculty(userDict["faculty"])
            roomNumber = RoomNumber(userDict["roomNumber"])
            excelDir = ExcelDir(userDict["userDict"])
            return User(studentId, userName, faculty, roomNumber, excelDir)

        else:
            return None

    def save(self, user: User):
        conn = sqlite3.connect("sample.db")
        c = conn.cursor()
        
