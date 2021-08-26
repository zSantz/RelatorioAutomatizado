import pandas as pd
import win32com.client as win32

#  Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

#  Visualizar a base de dados
pd.set_option('display.max_columns', None)

#  Faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()

#  Quantidade de produtos vendidos por loja
produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

#  Ticket médio por produto em cada loja
ticket_medio =  (faturamento['Valor Final'] / produtos_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})

#   Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'emailexemplo123@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de Produtos Vendidos:</p>
{produtos_vendidos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,<p>
<p>Equipe</p>

'''

mail.Send()