import sys
from cx_Freeze import setup, Executable

build_exe_options = {
    "packages": [
        "os",
        "tkinter",
        "customtkinter",
        "pandas",
        "matplotlib",
        "reportlab",
        "openpyxl"
    ],
    "includes": [
        "tkinter",
        "customtkinter",
        "ctkmessagebox"
    ],
    "include_files": [
        ("lojas_cadastradas.csv", "lojas_cadastradas.csv"),
        ("industrias_cadastradas.csv", "industrias_cadastradas.csv")
    ]
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Controle de Cupons",
    version="1.0",
    description="Controle de Cupons - Rede Sanar",
    options={"build_exe": build_exe_options, "include_msvcr": True},
    executables=[Executable("controle_cupons_sanar.py", base=base)]
)
