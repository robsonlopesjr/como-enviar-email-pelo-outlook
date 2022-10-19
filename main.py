import pandas as pd
import win32com.client as win32

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


def enviar_email(mail_to, faturamento, quantidade, ticket_medio):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = mail_to
    mail.Subject = 'Relatório de vendas por loja'
    mail.HTMLBody = f'''
    <p>Prezados,</p>
    <p>Segue o Relatório de Vendas por cada loja.</p>
    <p>Faturamento:</p>
    {faturamento.to_html()}
    <p>Quantidade Vendida:</p>
    {quantidade.to_html()}
    <p>Ticket Médio dos Produtos em cada loja:</p>
    {ticket_medio.to_html()}
    <p>Qualquer dúvida estou a disposição.</p>
    <p>Att.,</p>
    <p>Robson</p>
    '''

    mail.Send()

    print('E-mail Enviado.')

    return None


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

# Enviar um e-mail com o relatorio
enviar_email('teste@teste.com',
             faturamento, quantidade, ticket_medio)
