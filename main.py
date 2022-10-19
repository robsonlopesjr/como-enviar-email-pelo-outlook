import pandas as pd

# Configurações iniciais
pd.set_option('display.max_columns', None)  # Mostrar todas as colunas

# importar dados
dados = pd.read_excel('Vendas.xlsx')

# visualizar base de dados
# print(dados.head())

# Faturamento por loja
faturamento = dados[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja


# Ticket médio por produto em cada loja (Fatuamento / quantidade)


# Enviar um e-mail com o relatorio
