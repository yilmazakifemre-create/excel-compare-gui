@echo off
pip install pyinstaller
pyinstaller --onefile --noconsole main.py --icon=icon.ico --name="ExcelKarsilastirma"
pause
