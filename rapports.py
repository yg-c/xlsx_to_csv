import xlrd

path_in = input('File path: ')

print('lecture du fichier')

flag_start = -1
for i in range(len(path_in)):
    if path_in[i] == '\\':
        flag_start = i

flag_end = -1
for i in range(len(path_in)):
    if path_in[i] == '.':
        flag_end = i

file_name = path_in[flag_start + 1:flag_end]

path_out = path_in[:flag_start+1] + file_name + '.csv'

workbook = xlrd.open_workbook(path_in)
worksheet = workbook.sheet_by_index(0)
first_row = worksheet.row(0)
first_cell = worksheet.cell_value(0, 0)

if first_cell == 1839723408:
    file = open(path_out, 'w')
    line1 = first_cell
    file.writelines(str(line1) + "\n")
    file.close()
    print(file_name + ' --> écrit')
else:
    print(file_name + ' --> non conforme')

print('traitement terminé')

input()

# C:\Users\yannc\Desktop\tmp\test.xlsx
