from cx_Freeze import setup, Executable

base = "Win32GUI"   

executables = [Executable("rekap_nilai.py", base=base, icon="icon1.ico")]

packages = ["idna","PySimpleGUI","openpyxl","os","subprocess","win32com.client"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "Rekap Nilai Praktikan",
    options = options,
    version = "1.0.0",
    description = 'dibuat oleh Zaid Immaduddin Abdurrahman-202111129',
    executables = executables
)
# execute in folder prompt python setup.py build