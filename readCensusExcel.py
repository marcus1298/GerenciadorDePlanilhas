# cria uma tabela com a população e o número de setores censitários 
# de cada condado

import openpyxl, pprint
print('Opening workbook...')
wb = openpyxl.load_workbook('censuspopdata.xlsx')
sheet = wb.get_sheet_by_name('Population by Census Tract')
countyData = {}
# preenche countyData com a população e os setores de cada condado.
print('Reading rows...')
for row in range(2, sheet.max_row):
    #  Cada linha da planilha contém dados de um setor censitário.
    state  = sheet['B' + str(row)].value
    county = sheet['C' + str(row)].value
    pop    = sheet['D' + str(row)].value

    # Garante que a chave para esse estado existe
    countyData.setdefault(state, {})
    # Garante que a chave para esse condado nesse estado existe.
    countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})

    #  Cada linha representa um setor censitário, portanto incrementa o valor de um.
    countyData[state][county]['tracts'] += 1
    # Soma a população desse setor censitário à população do condado.
    countyData[state][county]['pop'] += int(pop)

# Abre um novo arquivo-texto e grava o conteúdo de countyData nesse arquivo.
print('Writing results...')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print('Done.')
