from openpyxl import load_workbook

from calc_kot1m import calc_kot1m

wb = load_workbook("Book1.xlsm")

sht_kot1 = wb["Котел I черга"]
sht_tur1 = wb["Турбіна I черга"]
sht_tur2 = wb["Турбіна II черга"]

calc_kot1m(sht_kot1, sht_tur1, sht_tur2)