from pathlib import Path
import sys


def abstractPath(relativePath):
    if getattr(sys, "frozen", False):
        basedir = Path(sys.executable).parent
    else:
        basedir = Path(__file__).parent
    return Path(basedir) / relativePath

UP = "▲"
DOWN = "▼"
SETTING_PATH = abstractPath("../user_setting.cfg")



