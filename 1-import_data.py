import pandas as pd

data = pd.read_excel("data/VendaCarros.xlsx")

# Lista os 5 primeiros
print(data.head())

# Lista os 5 últimos
print(data.tail())

# Listagem Específica
print(data["Fabricante"].value_counts())