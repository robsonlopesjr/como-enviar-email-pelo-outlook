import pandas as pd

# Configurações iniciais
pd.set_option('display.max_columns', None)  # Mostrar todas as colunas


def importar_dados(path):
    return pd.read_excel(path)


def visualizar_dados(dados):
    print(dados.head())
    return None


def calcular_faturamento(dados):
    return dados[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()


def calcular_quantidade_vendida(dados):
    return dados[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()


def calcular_ticket_medio(faturamento, quantidade):
    ticket = faturamento['Valor Final'] / quantidade['Quantidade']

    # transformar em tabela
    return ticket.to_frame()


# importar dados
arquivo = 'Vendas.xlsx'
dados = importar_dados(arquivo)

# visualizar base de dados
# visualizar_dados(dados)

# Faturamento por loja
faturamento = calcular_faturamento(dados)

# Quantidade de produtos vendidos por loja
quantidade = calcular_quantidade_vendida(dados)

# Ticket médio por produto em cada loja (Fatuamento / quantidade)
ticket_medio = calcular_ticket_medio(faturamento, quantidade)
visualizar_dados(ticket_medio)

# Enviar um e-mail com o relatorio
