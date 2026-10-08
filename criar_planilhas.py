import pandas as pd

vendas_sistema = pd.DataFrame({
    "ID": ["001", "002", "003", "004", "005"],
    "Cliente": ["João", "Maria", "Pedro", "Ana", "Lucas"],
    "Produto": ["Notebook", "Mouse", "Teclado", "Monitor", "Headset"],
    "Quantidade": [1, 2, 1, 1, 1],
    "Valor": [3500, 150, 200, 1200, 300],
    "Data": ["01/10/2026", "02/10/2026", "03/10/2026", "04/10/2026", "05/10/2026"]
})

vendas_financeiro = pd.DataFrame({
    "ID": ["001 ", "002", "003", "005", "006"],
    "Produto": ["Notebook", "Mouse", "Teclado", "Headset", "Webcam"],
    "Cliente": ["João Silva", "Maria", "Pedro", "Lucas", "Carlos"],
    "Valor": [3500, 180, 200, 300, 250],
    "Quantidade": [1, 3, 1, 1, 1],
    "Data": ["02/10/2026", "02/10/2026", "03/10/2026", "05/10/2026", "06/10/2026"]
})

vendas_sistema.to_excel("vendas_sistema.xlsx", index=False)
vendas_financeiro.to_excel("vendas_financeiro.xlsx", index=False)

print("Planilhas de teste criadas!")