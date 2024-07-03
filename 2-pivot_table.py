import pandas as pd

# Importando dados
data = pd.read_excel("data/VendaCarros.xlsx")
# print(type(data))

# Selecionando colunas específicas
df = data[["Fabricante", "ValorVenda", "Ano"]]
# print(df)

# Criando tabela pivô
pivot_table = df.pivot_table(
    index="Ano",
    columns="Fabricante",
    values="ValorVenda",
    aggfunc="sum"
)
# print(pivot_table)

# Exportando tabela pivô em arquivo excel
pivot_table.to_excel("data/pivot_table.xlsx", "Relatório")