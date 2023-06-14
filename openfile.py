import os
from pathlib import Path

string="открой word c названием документ"
idx=string.find("названием")
second_half = string[idx + len("названием"):].strip()
desktop_path = Path.home() / "Desktop"
file=f"{second_half}.doc"
file_path = os.path.join(desktop_path, file)
os.startfile(file_path)