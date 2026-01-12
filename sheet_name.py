import pandas as pd

# Lista as abas do Excel
print(pd.ExcelFile("arquivos\\2025_PROVA_TESTE.xlsx").sheet_names)
