import pandas as pd
import win32com.client as win32

# 1. importar a base de dados
# Ao importar fazer leitura no terminal com o seguinte comando: pip install openpyxl instalar e abrir
# arquivos python e excel)

tabela_vendas = pd.read_excel('Vendas.xlsx')

# 2. visualizar a base de dados
# Para tabela vendas abrir com o maximo de colunas --> display.max_colums

pd.set_option('display.max_columns', None)
print(tabela_vendas)

# 3. Faturamento por loja

# Usa-se Colchetes para filtar colunas de uma planilha podendo ser excel ou outra, ex:
# 2 colunas [[ ]] 2 colchetes, 3 colunas [[[ ]]] e assinm sucessivamente
# Para agrupar ons nomes ID usa -se > groupby
# Para somar usa- se --> sum
# Para contar usa- se --> count

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Separa as colunas
print('-' * 50)

# 4. Quantidade de produtos vendidos por loja

# Usa-se Colchetes para filtar colunas de uma planilha podendo ser excel ou outra, ex:
# 2 colunas [[ ]] 2 colchetes, 3 colunas [[[ ]]] e assinm sucessivamente
# Para agrupar ons nomes ID usa -se > groupby
# Para somar usa- se --> sum
# Para contar usa- se --> count

qtd_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Para alterar o nome dentro da coluna fica assim como no exemplo de Quantidade para Qtd_Produtos:
qtd_produtos = qtd_produtos.rename(columns={'Quantidade': 'Qtd_Produtos'})
print(qtd_produtos)

# Separa as colunas
print('-' * 50)

# 5. Ticket médio por produto em cada loja --> Divide o Faturamento(Valor Final / Qtd Produtos)

# Usa-se Colchetes para filtar colunas de uma planilha podendo ser excel ou outra, ex:
# 2 colunas [[ ]] 2 colchetes, 3 colunas [[[ ]]] e assinm sucessivamente
# Para agrupar ons nomes ID usa -se > groupby
# Para somar usa- se --> sum
# Para contar usa- se --> count
# Sempre que for fazer uma conta de uma Coluna para outra seja, soma, subtração, multiplicação ou divisão fica assim:
# ex --> ticket_medio = faturamento['Valor Final'] / qtd_produtos['Quantidade']
# Para deixar a tabela mais bonita coloca após o =  no começo o  abre ( e fecha após os nomes que tão dentro
# da tabela ) e finaliza com --> .to_frame()

ticket_medio = (faturamento['Valor Final'] / qtd_produtos['Qtd_Produtos']).to_frame()

# Para alterar o nome dentro da coluna fica assim como no exemplo de 0 para Ticket Medio:
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Separa as colunas
print('-' * 50)

# ¨6. Enviar um email com o relatório

# usa- se uma biblioteca que vai instalar no terminal de comando: pip install pywin32 --> (instalar e abrir
# arquivo python e Windows)
# Ao escrever o email usa HTML
# Para o texto ficar mais bonito usa- se .to_html e deixa o valor mais bonito também
# Ex: {tabela do faturamento.to_html(formatters={'Valor Final': 'R${:,.2f }'.format})}
# 2f --> significa  2 casas decimais}

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'rafaelraizer76@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'

# Aspas Simples ' ' o python escreve 1 texto.
# + Aspas Simples ''' ''' o python escreve mais linhas ou textos.

mail.HTMLBody = f'''  

<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

 <h2><b>Faturamento:</h2>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<h2><b>Quantidade Vendida:</h2>
{qtd_produtos.to_html(formatters={'Qtd_Produtos': 'R${:,.2f}'.format})}

<h2><b>Ticket Médio dos Produtos Vendidos:</h2>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Rafael Raizer</p>
'''

mail.Send()

print('Email Enviado')