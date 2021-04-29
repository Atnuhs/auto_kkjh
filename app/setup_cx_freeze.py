import sys
from cx_Freeze import setup, Executable


base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="test",
    version="0.1",
    description="My GUI application!",
    executables=[Executable("test.py", base=base)],
)
