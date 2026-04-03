import chardet

dat_file = "C:\\Users\\Robot\\Documents\\TEC\\PVT\\Py_calculation\\август 2015 норма.dat"
with open(dat_file, 'rb') as f:
    result = chardet.detect(f.read())

print(result)