import argparse
import xlrd

# parser = argparse.ArgumentParser()
# parser.add_argument('-in', dest='path_in', help='nom et extension du fichier')
# args = parser.parse_args()
# path_in = args.path_in

path = input("Chemin du fichier: ")

workbook = xlrd.open_workbook(path)
worksheet = workbook.sheet_by_index(0)
first_row = worksheet.row(0)
first_cell = worksheet.cell_value(0, 0)
print(first_cell)

if first_cell == 'GF':
    print('process')
    file = open('C:\\Users\\yannc\\Desktop\\tmp\\w1.csv', 'w')
    line1 = first_cell
    file.writelines(str(line1) + "\n")
    file.close()
    print('csv writed')
else:
    print('wrong file')

input()

# Windows cmd line :
# pip install xlrd
# C:\Users\yancla\PycharmProjects\rapport_de_caisse>py main.py -in test1.xlsx
