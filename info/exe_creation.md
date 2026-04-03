pip install pyinstaller

command to create exe:
pyinstaller --onedir --windowed --collect-all openpyxl --collect-all tkinter --hidden-import=openpyxl.cell._writer --name="Калькулятор_Котлів" main.py
