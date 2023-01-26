import sys
from cx_Freeze import setup, Executable

build_exe_options = {"packages": ['os', "playwright"],
                     "includes": ["tkinter", "pandas", "playwright", "tkintertable", "openpyxl", "os"],
                     }

base = None
if sys.platform == "win32":
    base = "win32GUI"

setup(
    name="PrintRessarcimento",
    version="0.1",
    description="Ressarcimento - Automação para fazer print de tela",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base)]
)
