import calendar
import datetime
import json
import subprocess
import sys
import typing as tp
from dataclasses import dataclass, asdict
from json import dump as jsondump
from json import load as jsonload
from pathlib import Path

import openpyxl
import PySimpleGUI as sg


def abstractPath(relativePath):
    if getattr(sys, "frozen", False):
        basedir = Path(sys.excutable).parent
    else:
        basedir = Path(__file__).parent
    return Path(basedir) / relativePath


UP = "▲"
DOWN = "▼"
SETTING_PATH = abstractPath("../user_setting.cfg")
EXCEL_TEMPLATE_PATH = abstractPath("template/template.xlsx")
DEFAULT_SETTING = {
    "faculty": "自然科学研究科数理物質専攻",
    "studentID": "",
    "userName": "",
    "roomNumber": "",
    "excelFileLocation": "",
}


@dataclass
class UserSetting:
    faculty: str
    studentID: str
    userName: str
    roomNumber: str
    excelFileLocation: str

    def save_to_json(self, dst: Path):
        with open(dst, "w") as f:
            jsondump(asdict(self), f)

    @staticmethod
    def load_from_json(dst: Path = SETTING_PATH):
        try:
            with open(dst, "r") as f:
                userSettingDict = jsonload(f)
        except FileNotFoundError:
            userSettingDict = DEFAULT_SETTING
        except json.decoder.JSONDecodeError:
            dst.unlink()
            userSettingDict = DEFAULT_SETTING

        try:
            userSetting = UserSetting(**userSettingDict)
        except TypeError:
            dst.unlink()
            userSettingDict = DEFAULT_SETTING
            userSetting = UserSetting(**userSettingDict)
        return userSetting


class AttendanceRecord:
    @staticmethod
    def excelFileName():
        today = datetime.date.today()
        userName = UserSetting.load_from_json().userName
        return f"研究活動状況表(R{today.year-2018}.{today.month})_{userName}.xlsx"

    def excelPath(self):
        loc = UserSetting.load_from_json().excelFileLocation
        return Path(loc) / self.excelFileName()

    def newExcelFromTemplate(self) -> openpyxl.workbook:
        today = datetime.date.today()
        workbook = openpyxl.load_workbook(EXCEL_TEMPLATE_PATH)
        sheet = workbook.active
        sheet.title = f"{today.month}月"
        sheet["A2"] = f"令和{today.year-2018}年{today.month}月　研究活動状況表（学生用）"
        num_day = calendar.monthrange(today.year, today.month)[1]
        for i in range(num_day):
            cell = f"A{i+16}"
            day = i + 1
            ymd = datetime.date(today.year, today.month, day)
            sheet[cell] = ymd
            sheet[cell].number_format = "m月d日"
        return workbook

    def getExcel(self) -> openpyxl.workbook:
        target = self.excelPath()
        if target.is_file():
            workbook = openpyxl.load_workbook(target)
        else:
            workbook = self.newExcelFromTemplate()
        return workbook

    def saveExcel(self, workbook: openpyxl.workbook):
        workbook.save(self.excelPath())

    def stamp_roomNumber(self):
        workbook = self.getExcel()
        sheet = workbook.active
        dt_now = datetime.datetime.now()
        cell = f"H{dt_now.day+15}"
        newVal = self.userSetting.roomNumber
        print(f"set roomNumber => cell: {cell} roomNumber: {newVal}")
        sheet[cell] = newVal
        self.saveExcel(workbook)

    def stamp_entry_time(self):
        workbook = self.getExcel()
        sheet = workbook.active
        dt_now = datetime.datetime.now()
        cell = f"C{dt_now.day+15}"
        val = dt_now.time()
        print(f"stamp entry_time => cell:{cell} time:{val}")
        sheet[cell] = val
        self.saveExcel(workbook)

    def stamp_exit_time(self):
        workbook = self.getExcel()
        sheet = workbook.active
        dt_now = datetime.datetime.now()
        cell = f"E{dt_now.day+15}"
        val = dt_now.time()
        print(f"stamp exit_time => cell:{cell} time:{val}")
        sheet[cell] = val
        self.saveExcel(workbook)

    def today_entry_time(self):
        sheet = self.getExcel().active
        dt_now = datetime.datetime.now()
        time = sheet[f"C{dt_now.day+15}"].value
        return datetime.datetime.combine(dt_now.date(), time) if time else None

    def today_exit_time(self):
        sheet = self.getExcel().active
        dt_now = datetime.datetime.now()
        time = sheet[f"E{dt_now.day+15}"].value
        return datetime.datetime.combine(dt_now.date(), time) if time else None


class Mainwindow:
    def __init__(self):
        self.open = False
        button_size = (10, 2)
        entry_text = sg.Text(
            size=button_size, justification="center", key="-ENTRY_TEXT-"
        )
        exit_text = sg.Text(size=button_size, justification="center", key="-EXIT_TEXT-")
        open_sec = sg.Text(
            UP, enable_events=True, key="-OPEN_SEC-", text_color="yellow"
        )
        today_text = sg.Text(size=(25, 2), key="-TODAY_TEXT-")
        user_data_text = sg.Text(size=(40, 6), key="-USER_DATA_TEXT-")
        section = [
            [entry_text, exit_text],
            [
                sg.Button("設定", size=button_size),
                sg.Button("使い方", size=button_size),
                sg.Button("Excelで開く", size=button_size),
            ],
            [today_text],
            [sg.Text("ユーザ情報")],
            [user_data_text],
        ]
        layout = [
            [sg.Button("入室", size=button_size), sg.Button("退室", size=button_size)],
            [open_sec],
            [collapse(section, "-SEC-", self.open)],
        ]
        self.window = sg.Window("auto_kkjh", layout)

    def show_window(self):
        return self.window.read(timeout=2000, timeout_key="-TIMEOUT-")

    def toggle_sec(self):
        self.open = not self.open
        self.window["-OPEN_SEC-"].update(DOWN if self.open else UP)
        self.window["-SEC-"].update(visible=self.open)

    def time_update(self, entry_time, exit_time):
        if entry_time == None and exit_time == None:
            update_text = "おはよう"
        else:
            if entry_time:
                today_time = datetime.datetime.now() - entry_time
                snt = "入室"
            if exit_time:
                today_time = datetime.datetime.now() - exit_time
                snt = "退室"
            h, m, _ = get_h_m_s(today_time)
            update_text = f"{snt}してから{h}時間{m}分経過"
        self.window["-TODAY_TEXT-"].update(update_text)

    def update_entry_time(self, entry_time):
        update_text = entry_time.strftime("%H:%M 入室") if entry_time else "まだ入室してない"
        self.window["-ENTRY_TEXT-"].update(update_text)

    def update_exit_time(self, exit_time):
        update_text = exit_time.strftime("%H:%M 退室") if exit_time else "まだ退室してない"
        self.window["-EXIT_TEXT-"].update(update_text)

    def update_user_data(self, us):
        update_text = f"名前: {us.userName}\n学籍番号: {us.studentID}\n所属: {us.faculty}\n部屋番号 {us.roomNumber}"
        print(update_text)
        self.window["-USER_DATA_TEXT-"].update(update_text)


class SettingWindow:
    def show_window(self, usersetting: tp.Type[UserSetting]) -> tp.Type[UserSetting]:
        text_size = (15, 1)
        text_pad = ((10, 10), (10, 10))
        button_size = (10, 1)
        button_pad = ((20, 20), (10, 10))
        layout = [
            [sg.Text("現在のユーザ情報", size=text_size, pad=text_pad)],
            [
                sg.Text("学部・大学院", size=text_size, pad=text_pad),
                sg.Input(usersetting.faculty, key="faculty"),
            ],
            [
                sg.Text("学籍番号", size=text_size, pad=text_pad),
                sg.Input(usersetting.studentID, key="studentID"),
            ],
            [
                sg.Text("名前", size=text_size, pad=text_pad),
                sg.Input(usersetting.userName, key="userName"),
            ],
            [
                sg.Text("部屋番号", size=text_size, pad=text_pad),
                sg.Input(usersetting.roomNumber, key="roomNumber"),
            ],
            [
                sg.Text("エクセルファイルの場所", size=text_size, pad=text_pad),
                sg.Input(usersetting.excelFileLocation, key="excelFileLocation"),
            ],
            [
                sg.Button("OK", size=button_size, pad=button_pad),
                sg.Button("キャンセル", size=button_size, pad=button_pad),
            ],
        ]
        window = sg.Window("auto_kkjh設定", layout)
        event, values = window.read()
        window.close()
        if event == sg.WINDOW_CLOSED or event == "キャンセル":
            newUserSetting = usersetting
        if event == "OK":
            newUserSetting = UserSetting(**values)
        return newUserSetting

    def save_usersetting(self, usersetting):
        with open(SETTING_PATH, "w") as f:
            jsondump(usersetting.to_json(), f)


class ManualWindow:
    def show_window(self):
        layout = [
            [sg.Text("入室したら入室ボタン、退室するときは退室ボタンを押す。")],
            [sg.Text("名前の変更、月が変わった時などは新しくExcelが自動生成されます")],
            [sg.Text("学部・大学院、学籍番号、部屋番号は変えても元のExcelに上書きされ、新規作成はされません。")],
            [sg.Button("了解!")],
        ]
        window = sg.Window("丁寧な使い方の説明を見て思わずあふれ出る涙", layout)
        window.read()
        window.close()


def collapse(layout, key, visible):
    return sg.pin(sg.Column(layout, key=key, visible=visible))


def get_h_m_s(td):
    m, s = divmod(td.seconds, 60)
    h, m = divmod(m, 60)
    return h, m, s


def load_setting(setting_file) -> tp.Type[UserSetting]:
    setting_filepath = Path(setting_file)
    if setting_filepath.is_file():
        try:
            with open(file=setting_filepath, mode="r") as f:
                usersetting = jsonload(f)
        except json.decoder.JSONDecodeError:
            setting_filepath.unlink()
            usersetting = DEFAULT_SETTING
    else:
        usersetting = DEFAULT_SETTING
    return UserSetting(**usersetting)


def main():
    mw = Mainwindow()
    sw = SettingWindow()
    manw = ManualWindow()
    if not Path(SETTING_PATH).is_file():
        usersetting = sw.show_window(UserSetting.load_from_json())
        usersetting.save_to_json(SETTING_PATH)
    ar = AttendanceRecord()

    while True:
        event, values = mw.show_window()

        if event == sg.WINDOW_CLOSED:
            break
        print(event)
        print(values)
        if event == "使い方":
            manw.show_window()
        if event == "-OPEN_SEC-":
            mw.toggle_sec()
        if event == "入室":
            ar.stamp_entry_time()
            mw.update_entry_time(ar.today_entry_time())
            mw.time_update(ar.today_entry_time(), ar.today_exit_time())
        if event == "退室":
            ar.stamp_exit_time()
            mw.update_exit_time(ar.today_exit_time())
            mw.time_update(ar.today_entry_time(), ar.today_exit_time())
        if event == "設定":
            usersetting = sw.show_window(UserSetting.load_from_json())
            usersetting.save_to_json(SETTING_PATH)
            mw.update_user_data(UserSetting.load_from_json())
        if event == "Excelで開く":
            subprocess.Popen(
                ["start", ar.excelFileName(UserSetting.load_from_json().userName)],
                shell=True,
            )
            break
        if event == "-TIMEOUT-":
            mw.update_entry_time(ar.today_entry_time())
            mw.update_exit_time(ar.today_exit_time())
            mw.time_update(ar.today_entry_time(), ar.today_exit_time())
            mw.update_user_data(UserSetting.load_from_json())
            continue

    mw.window.close()


if __name__ == "__main__":
    # os.chdir(Path(__file__).parent)
    main()
