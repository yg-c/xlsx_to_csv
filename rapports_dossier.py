import os
import xlrd
import getpass

# constants
mandant = '3963'
devise = 'CHF'
version = 'J'
quantite = '0'
separator = ','
end = '0,0,0,0,0,0,0,0,E'
empty = ''
dot = '.'
username = getpass.getuser()

print('-- Comptabilisation dossier rapport de caisse --' + '\n')
print('Tchôôôôôô ' + username)
folder = input('Chemin du dossier: ')
print('\n' + '-- lecture du dossier --' + '\n')
# directory control
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
        # file extension control
        if r[flag_extension:] == '.xlsx':
            path_out = folder + '\\' + r[:flag_extension] + '.csb'
            workbook = xlrd.open_workbook(folder + '\\' + r)
            worksheet = workbook.sheet_by_index(0)
            #  first_row = worksheet.row(0)
            first_cell = worksheet.cell_value(0, 0)  # row, column
            # file content control
            if first_cell == 1839723408:
                file = open(path_out, 'w', encoding='utf-8')
                # csb row number
                number = 1
                annee = str(int(worksheet.cell_value(1, 6)))
                mois = str(int(worksheet.cell_value(1, 7)))
                jour = str(int(worksheet.cell_value(1, 8)))
                caisse = str(int(worksheet.cell_value(3, 7)))
                numero_ecriture = str(worksheet.cell_value(1, 2))
                for i in range(4, 71):
                    # write row or not
                    if worksheet.cell_value(i, 12) == 1:
                        if worksheet.cell_value(i, 7) != '':
                            compte = str(int(worksheet.cell_value(i, 7)))
                        else:
                            compte = str(worksheet.cell_value(i, 7))

                        libelle = str(worksheet.cell_value(i, 5))
                        montant = str(worksheet.cell_value(i, 11))
                        debit_credit = str(worksheet.cell_value(i, 6))

                        if worksheet.cell_value(i, 10) != '':
                            section = str(int(worksheet.cell_value(i, 10)))
                        else:
                            section = str(worksheet.cell_value(i, 10))

                        if worksheet.cell_value(i, 8) != '':
                            code_tva = str(int(worksheet.cell_value(i, 8)))
                        else:
                            code_tva = str(worksheet.cell_value(i, 8))

                        taux_tva = str(worksheet.cell_value(i, 9))

                        if taux_tva != '':
                            tva_inclu = 'I'
                        else:
                            tva_inclu = ''

                        if taux_tva != '':
                            coeff_tva = '100'
                        else:
                            coeff_tva = ''

                        if taux_tva != '':
                            methode_tva = '1'
                        else:
                            methode_tva = '0'

                        file.writelines(str(number) + separator + version + separator + jour + dot + mois + dot + annee
                                        + separator + caisse + separator + compte + separator + libelle + separator
                                        + montant + separator + debit_credit + separator + separator + section
                                        + separator + numero_ecriture + separator + mandant + separator + devise
                                        + separator + devise + separator + quantite + separator + code_tva
                                        + separator + taux_tva + separator + tva_inclu + separator + methode_tva
                                        + separator + coeff_tva + separator + end + "\n")
                        number += 1

                file.close()
                print(r[:flag_extension] + '.csb' + ' --> écrit')
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

print('\n' + '-- <\> by yancla --' + '\n')
input()

# C:\Users\yannc\Desktop\tmp
