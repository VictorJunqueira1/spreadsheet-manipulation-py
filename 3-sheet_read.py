from openpyxl import load_workbook

# Lê a pasta de trabalho e planilha
wb = load_workbook("data/pivot_table.xlsx")
sheet = wb["Relatório"]

# Acessando um valor específico
# print(sheet["B3"].value) 

# Iterando valores por meio de loop
for i in range(2, 6):
    ano = sheet["A%s" %i].value
    astonMartin = sheet["B%s" %i].value
    bentley = sheet["C%s" %i].value
    print("{0} o Aston Martin vendeu {1}, e o Bentley vendeu {2}".format(ano, astonMartin, bentley))