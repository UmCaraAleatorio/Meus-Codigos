import pandas as pd  # este é para importar os arquivos em  excel
import win32com.client as win32  # este é para enviar emails para o outlook

# importar a base de dados

# com isso ele transfere o arquivo em excel para o python
tabela_vendas = pd.read_excel('C:/Users/Gamemax/Desktop/Projeto/Pasta do minicurso/Vendas.xlsx')


# visualizar a base de dados

# com esse ele tira o limete de colunas que o algoritimo mostra
pd.set_option('display.max_columns', None)

# faturamento por loja

# vai olhar as colunas que estão escritas e vai agrupar a(s) coluna(s) que estão escritas e vai somar as linhas destas colunas
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()

# quantidade de produtos vendidos por loja

# vai olhar as colunas que estão escritas e vai agrupar a(s) coluna(s) que estão escritas e vai somar as linhas destas colunas
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# ticket medio por produto em cada loja

# este vai calcular a media de cada loja e vai colocar em uma planilha corretamente
ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# enviar um email com o relatorio


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'Email vai aqui'
mail.Subject = 'Relatorio de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados</p>

<p>Segue o Relatorio de vendas por cada loja:</p

<p>Faturamento:</p>

{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>

{quantidade.to_html()}

<p>Ticket médio dos Produtos em cada loja:</p>

{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou à disposição.</p>

<p>Att.,</p>
<p>Victor</p>
'''

mail.Send()
