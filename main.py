import pandas as pd

# importar dados
dados = pd.read_excel('Vendas.xlsx')

# visualizar base de dados
pd.set_option('display.max_columns', None)  # Mostrar todas as colunas
print(dados.head())

# Faturamento por loja


# Quantidade de produtos vendidos por loja


# Ticket médio por produto em cada loja (Fatuamento / quantidade)


# Enviar um e-mail com o relatorio
