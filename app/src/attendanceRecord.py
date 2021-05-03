import calendar
import openpyxl
from pathlib import Path
from datetime import date, datetime
from util import abstractPath


EXCEL_TEMPLATE_PATH = abstractPath("template/template.xlsx")


class AttendanceRecord:
    @staticmethod
    def newExcelFromTemplate(self) -> openpyxl.workbook:
        print("Generate New excell from template")
        today = date.today()
        workbook = openpyxl.load_workbook(EXCEL_TEMPLATE_PATH)
        sheet = workbook.active
        sheet.title = f"{today.month}月"
        sheet["A2"] = f"令和{today.year-2018}年{today.month}月　研究活動状況表（学生用）"
        num_day = calendar.monthrange(today.year, today.month)[1]
        for i in range(num_day):
            cell = f"A{i+16}"
            day = i + 1
            ymd = date(today.year, today.month, day)
            sheet[cell] = ymd
            sheet[cell].number_format = "m月d日"
        return workbook

    def getExcel(self, excelPath: Path) -> openpyxl.workbook:
        print("Open excel")
        if excelPath.is_file():
            return openpyxl.load_workbook(excelPath)
        return self.newExcelFromTemplate()

    def saveExcel(self, workbook: openpyxl.workbook, excelPath: Path):
        print("save excel")
        workbook.save(excelPath)

    def stampRoomNumber(self, excelPath: Path):
        dt_now = datetime.now()
        cell = f"H{dt_now.day+15}"
        newVal = self.userSetting.roomNumber
        print(f"set roomNumber => cell: {cell} roomNumber: {newVal}")
        workbook = self.getExcel(excelPath)
        sheet = workbook.active
        sheet[cell] = newVal
        self.saveExcel(workbook, userName)

    def stampUserSetting(self, userSetting: User):
        print(f"stamp faculty => {userSetting.faculty}")
        print(f"stamp studentID => {userSetting.studentID}")
        print(f"stamp username => {userSetting.userName}")
        workbook = self.getExcel(userSetting.userName)
        sheet = workbook.active
        sheet["G7"] = userSetting.faculty
        sheet["G8"] = userSetting.studentID
        sheet["G9"] = userSetting.userName
        self.saveExcel(workbook, userSetting)

    def stampEntryTime(self, userName: str):
        dt_now = datetime.now()
        cell = f"C{dt_now.day+15}"
        val = dt_now.time()
        print(f"stamp entry_time => cell:{cell} time:{val}")
        workbook = self.getExcel(userName)
        sheet = workbook.active
        sheet[cell] = val
        self.saveExcel(workbook, userName)

    def stampExitTime(self, userName: str):
        dt_now = datetime.now()
        cell = f"E{dt_now.day+15}"
        val = dt_now.time()
        print(f"stamp exit_time => cell:{cell} time:{val}")
        workbook = self.getExcel(userName)
        sheet = workbook.active
        sheet[cell] = val
        self.saveExcel(workbook, userName)

    def TodayEntryTime(self, userName: str):
        print("See today's entry time")
        sheet = self.getExcel(userName).active
        dt_now = datetime.now()
        time = sheet[f"C{dt_now.day+15}"].value
        return datetime.combine(dt_now.date(), time) if time else None

    def TodayExitTime(self, userName: str):
        print("See today's exit time")
        sheet = self.getExcel(userName).active
        dt_now = datetime.now()
        time = sheet[f"E{dt_now.day+15}"].value
        return datetime.combine(dt_now.date(), time) if time else None
