import pandas as pd

# Importar a base de dados
tabelas_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabelas_vendas)

print('-' * 50)
# Calcular o Faturamento
faturamento = tabelas_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('-' * 50)
# Qtd de Produtos vendidos por loja
quantidade = tabelas_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
print(ticket_medio)

# Enviar um email com o relatorio
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'caiozao1212@gmail.com'
mail.Subject = 'Relatorio de Vendas por loja'
mail.HTMLBody = '''
<p>Preazados, </p>

<p>Segue o Relatório de Vendas por Cada Loja: </p>

<p>Faturamento: </p>
{faturamento.to_html()}

<p>Quantidade Vendida: </p>
{quantidade.to_html()}

<p>Ticket Medio: </p>
{ticket_medio.to_html()}


<p>Qualquer duvida, estou a disposição. </p>
'''

mail.Send()
print("Email Enviado.")

