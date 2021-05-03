from datetime import date
from dataclasses import dataclass, asdict
import json

DEFAULT_SETTING = {
    "faculty": "自然科学研究科数理物質専攻",
    "studentID": "",
    "userName": "",
    "roomNumber": "",
    "excelFileLocation": "",
}

class UserSettingService:
    pass

class UserSettingRegisory:
    pass

@dataclass
class UserSetting:
    studentID: str
    userName: str
    faculty: str
    roomNumber: str
    excelDir: str

    def excelFileName(self):
        today = date.today()
        fileSuffix = f"(R{today.year-2018}.{today.month})_{self.userName}.xlsx"
        return f"研究活動状況表{fileSuffix}"

    def excelPath(self):
        loc = self.excelDir
        return Path(loc) / self.excelFileName()

    def save_to_json(self, dst: Path):
        print(f"save setting {dst}")
        with open(dst, "w") as f:
            json.dump(asdict(self), f)

    @staticmethod
    def load_from_json(dst: Path = SETTING_PATH):
        print(f"< load setting {dst}")
        try:
            with open(dst, "r") as f:
                userSettingDict = jsonload(f)
        except FileNotFoundError:
            userSettingDict = DEFAULT_SETTING
            print(f"setting file {dst} not found")
        except json.decoder.JSONDecodeError:
            dst.unlink()
            userSettingDict = DEFAULT_SETTING
            print(f"setting file {dst} is broken")

        try:
            userSetting = UserSetting(**userSettingDict)
        except TypeError:
            dst.unlink()
            userSettingDict = DEFAULT_SETTING
            userSetting = UserSetting(**userSettingDict)
            print("userSetting broken >")
        return userSetting
