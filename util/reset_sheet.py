from openpyxl import load_workbook

from reset_computed_cells import reset_computed_zagal_cells

wb = load_workbook("G:\\other\\PTV\\Py_calculation\\data\\exel\\cerkassy_test.xlsx")

sht = wb["Загальні"]

for row in sht["A1":"D10"]:  # наприклад, від A1 до D10
    for cell in row:
        print(cell.value, end="\t")
    print()

reset_computed_zagal_cells(sht)

wb.save("G:\\other\\PTV\\Py_calculation\\data\\exel\\cerkassy_test_zagal.xlsx")
