import os
import subprocess
from pathlib import Path
import nicegui

cmd = [
    'python',
    '-m', 'PyInstaller',
    '../src/main.py',  # your main file with ui.run()
    '--name', 'ExcelWithPic',  # name of your app
    '--icon', '../favicon.ico',
    '--onefile',
    '--clean',
    '--windowed',  # prevent console appearing, only use with ui.run(native=True, ...)
    '--add-data', f'{Path(nicegui.__file__).parent}{os.pathsep}nicegui',
    '--hidden-import=openpyxl.cell._writer'
]
subprocess.call(cmd)
