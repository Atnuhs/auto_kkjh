import calendar
import datetime
import os
import sys
import subprocess
from json import dump as jsondump
from json import load as jsonload
import json
from pathlib import Path

import openpyxl
import PySimpleGUI as sg
from openpyxl.styles import Font


class AttendanceRecord:
    def year(self):
        return datetime.date.today().year

    def reiwa_year(self):
        return self.year() - 2018

    def month(self):
        return datetime.date.today().month

    def day(self):
        return datetime.date.today().day

    def hour(self):
        return datetime.datetime.now().hour

    def minute(self):
        return datetime.datetime.now().minute

    def xlname(self, username):
        return f"研究活動状況表(R{self.reiwa_year()}.{self.month()})_{username}.xlsx"

    def save_record(self, usersetting):
        print(f"set faculty => {usersetting.faculty}")
        print(f"set studentID => {usersetting.studentID}")
        print(f"set username => {usersetting.username}")
        self.sheet["G7"] = usersetting.faculty
        self.sheet["G8"] = usersetting.studentID
        self.sheet["G9"] = usersetting.username
        self.wb.save(self.xlname(usersetting.username))

    def generate_wb_from_template(self):
        wb = openpyxl.load_workbook(EXCEL_TEMPLATE_FILE)
        sheet = wb.active
        sheet.title = f"{self.month()}月"
        sheet["A2"] = f"令和{self.reiwa_year()}年{self.month()}月　研究活動状況表（学生用）"
        num_day = calendar.monthrange(self.year(), self.month())[1]
        for i in range(num_day):
            cell = f"A{i+16}"
            ymd = datetime.date(self.year(), self.month(), i + 1)
            sheet[cell] = ymd
            sheet[cell].number_format = "m月d日"
        return wb

    def stamp_roomnumber(self, roomnumber):
        cell = f"H{self.day()+15}"
        print(f"set roomnumber => cell: {cell} roomnumber: {roomnumber}")
        self.sheet[cell] = roomnumber

    def stamp_entry_time(self):
        cell = f"C{15+self.day()}"
        time = datetime.datetime.now().time()
        print(f"stamp entry_time => cell:{cell} time:{time}")
        self.sheet[cell] = time

    def stamp_exit_time(self):
        cell = f"E{15+self.day()}"
        time = datetime.datetime.now().time()
        print(f"stamp exit_time => cell:{cell} time:{time}")
        self.sheet[cell] = time

    def today_entry_time(self):
        cell = self.sheet[f"C{15+self.day()}"].value
        return datetime.datetime.combine(datetime.date.today(), cell) if cell else None

    def today_exit_time(self):
        cell = self.sheet[f"E{15+self.day()}"].value
        return datetime.datetime.combine(datetime.date.today(), cell) if cell else None


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
            date = datetime.date.today()
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
        update_text = f"名前: {us.username}\n学籍番号: {us.studentID}\n所属: {us.faculty}\n部屋番号 {us.roomnumber}"
        print(update_text)
        self.window["-USER_DATA_TEXT-"].update(update_text)


class SettingWindow:
    def show_window(self, usersetting):
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
                sg.Input(usersetting.username, key="username"),
            ],
            [
                sg.Text("部屋番号", size=text_size, pad=text_pad),
                sg.Input(usersetting.roomnumber, key="roomnumber"),
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
            return usersetting
        if event == "OK":
            self.save_usersetting(UserSetting(**values))
        return UserSetting(**values)

    def save_usersetting(self, usersetting):
        with open(SETTING_FILE, "w") as f:
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


class UserSetting:
    def __init__(self, faculty, studentID, username, roomnumber):
        self.faculty = faculty
        self.studentID = studentID
        self.username = username
        self.roomnumber = roomnumber

    def to_json(self):
        return {
            "faculty": self.faculty,
            "studentID": self.studentID,
            "username": self.username,
            "roomnumber": self.roomnumber,
        }


def collapse(layout, key, visible):
    return sg.pin(sg.Column(layout, key=key, visible=visible))


def get_h_m_s(td):
    m, s = divmod(td.seconds, 60)
    h, m = divmod(m, 60)
    return h, m, s


def load_setting(setting_file):
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


def load_excel(username):
    an = AttendanceRecord()
    if Path(an.xlname(username)).is_file():
        an.wb = openpyxl.load_workbook(an.xlname(username))
    else:
        an.wb = an.generate_wb_from_template()
    an.sheet = an.wb.active
    return an


def abstractPath(relativePath):
    if getattr(sys, "frozen", False):
        basedir = Path(sys.excutable).parent
    else:
        basedir = Path(__file__).parent
    return Path(basedir) / relativePath


UP = "▲"
DOWN = "▼"
SETTING_FILE = abstractPath("../user_setting.cfg")
EXCEL_TEMPLATE_FILE = abstractPath("template/template.xlsx")
DEFAULT_SETTING = {
    "faculty": "自然科学研究科数理物質専攻",
    "studentID": "",
    "username": "",
    "roomnumber": "",
}


def main():
    mw = Mainwindow()
    sw = SettingWindow()
    manw = ManualWindow()
    usersetting = load_setting(SETTING_FILE)
    if not Path(SETTING_FILE).is_file():
        usersetting = sw.show_window(usersetting)
    ar = load_excel(usersetting.username)

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
            usersetting = sw.show_window(usersetting)
            mw.update_user_data(usersetting)
            ar = load_excel(usersetting.username)
        if event == "Excelで開く":
            subprocess.Popen(["start", ar.xlname(usersetting.username)], shell=True)
            break
        if event == "-TIMEOUT-":
            mw.update_entry_time(ar.today_entry_time())
            mw.update_exit_time(ar.today_exit_time())
            mw.time_update(ar.today_entry_time(), ar.today_exit_time())
            mw.update_user_data(usersetting)
            continue
        ar.save_record(usersetting)

    mw.window.close()


if __name__ == "__main__":
    # os.chdir(Path(__file__).parent)
    main()
