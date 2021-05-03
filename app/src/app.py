import datetime
import subprocess
from json import dump as jsondump
import PySimpleGUI as sg


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
    def show_window(self, usersetting: User) -> User:
        def Line(string: str):
            text_size = (15, 1)
            text_pad = ((10, 10), (10, 10))
            return sg.T(string, size=text_size, pad=text_pad)

        button_size = (10, 1)
        button_pad = ((20, 20), (10, 10))
        layout = [
            [Line("現在のユーザ情報")],
            [
                Line("学部・大学院"),
                sg.I(usersetting.faculty, key="faculty"),
            ],
            [
                Line("学籍番号"),
                sg.I(usersetting.studentID, key="studentID"),
            ],
            [
                Line("名前"),
                sg.I(usersetting.userName, key="userName"),
            ],
            [
                Line("部屋番号"),
                sg.I(usersetting.roomNumber, key="roomNumber"),
            ],
            [
                Line("エクセル出力先"),
                sg.I(usersetting.excelFileLocation, key="excelFileLocation"),
                sg.FolderBrowse("フォルダの参照", target="excelFileLocation"),
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


def main():
    main_menu = Mainwindow()
    setting_menu = SettingWindow()
    manual_menu = ManualWindow()
    userSetting = UserSetting.load_from_json()
    if not Path(SETTING_PATH).is_file():
        userSetting = setting_menu.show_window(userSetting)
        userSetting.save_to_json(SETTING_PATH)
    ar = AttendanceRecord()

    while True:
        event, values = main_menu.show_window()

        if event == sg.WINDOW_CLOSED:
            break
        print(event)
        print(values)
        if event == "使い方":
            manual_menu.show_window()
        if event == "-OPEN_SEC-":
            main_menu.toggle_sec()
        if event == "入室":
            ar.stampEntryTime(userSetting.userName)
        if event == "退室":
            ar.stampExitTime(userSetting.userName)
        if event == "設定":
            userSetting = setting_menu.show_window(userSetting)
            userSetting.save_to_json(SETTING_PATH)
            main_menu.update_user_data(userSetting)
            ar.stampUserSetting(userSetting)
        if event == "Excelで開く":
            subprocess.Popen(
                ["start", ar.excelFileName(userSetting.userName)],
                shell=True,
            )
            break
        if event == "-TIMEOUT-":
            todayEntryTime = ar.TodayEntryTime(userSetting.userName)
            todayExitTime = ar.TodayExitTime(userSetting.userName)
            main_menu.update_entry_time(todayEntryTime)
            main_menu.update_exit_time(todayEntryTime)
            main_menu.time_update(todayEntryTime, todayExitTime)
            # main_menu.update_user_data(UserSetting.load_from_json())
            continue

    main_menu.window.close()


if __name__ == "__main__":
    # os.chdir(Path(__file__).parent)
    main()
