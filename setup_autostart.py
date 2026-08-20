"""Registers Daily Todo to launch silently at Windows login.

Run this once: python setup_autostart.py
"""
import os
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
APP_SCRIPT = APP_DIR / "todo_app.py"

pythonw = Path(sys.executable).with_name("pythonw.exe")
if not pythonw.exists():
    pythonw = Path(sys.executable)  # fallback, will show a console window

startup_dir = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
vbs_path = startup_dir / "DailyTodo.vbs"

vbs_content = (
    'Set WshShell = CreateObject("WScript.Shell")\n'
    f'WshShell.Run """{pythonw}"" ""{APP_SCRIPT}""", 0\n'
    'Set WshShell = Nothing\n'
)

startup_dir.mkdir(parents=True, exist_ok=True)
vbs_path.write_text(vbs_content, encoding="utf-8")

print(f"Autostart entry created: {vbs_path}")
print("Daily Todo will now launch silently (no console window) every time you log in.")
print("To undo: delete that file.")
