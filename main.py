import pandas as pd

# Importar a base de dados
tabelas_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabelas_vendas)

# Calcular o Faturamento
faturamento = tabelas_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Qtd de Produtos vendidos por loja

# Ticket médio por produto em cada loja
# Enviar um email com o relatorio
