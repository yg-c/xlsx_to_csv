import os
import xlrd

folder = input('Chemin du dossier: ')
print('\n' + '-- lecture du dossier --' + '\n')
if os.path.isdir(folder):

    arr = os.listdir(folder)

    print('***')
    for r in arr:
        print(r + ' --> traitement')

        file_name = r

        flag_extension = -1
        for i in range(len(r)):
            if r[i] == '.':
                flag_extension = i

        if r[flag_extension:] == '.xlsx':
            path_out = folder + '\\' + r[:flag_extension] + '.csv'
            workbook = xlrd.open_workbook(folder + '\\' + r)
            worksheet = workbook.sheet_by_index(0)
            first_row = worksheet.row(0)
            first_cell = worksheet.cell_value(0, 0)

            if first_cell == 1839723408:
                file = open(path_out, 'w')
                line1 = first_cell
                file.writelines(str(line1) + "\n")
                file.close()
                print(r[:flag_extension] + '.csv' + ' --> écrit')
                print('***')
            else:
                print(r + ' --> contenu non conforme')
                print('***')

        else:
            print(r + ' --> extension non conforme')
            print('***')

    print('\n' + '-- traitement terminé --')

else:
    print('-- dossier non valide --')

input()

# C:\Users\yannc\Desktop\tmp
