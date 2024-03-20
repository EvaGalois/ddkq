import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": ["tkinter", "pandas", "xlsxwriter", "openpyxl"],
    "include_files": [],  # 如果有额外的文件需要包含，请在这里添加
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"  # 如果是Windows GUI应用，使用"Win32GUI"

executables = [Executable("Attendance_new.py", base=base, target_name="工时计算机.exe", icon="app_icon.ico")]

setup(
    name="ddkq",
    version="1.0",
    description="ddkq",
    options={"build_exe": build_exe_options},
    executables=executables
)
