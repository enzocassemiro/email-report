import pandas as pd
import win32com.client as win32

# importando base de dados
tabelaVendas = pd.read_excel('Vendas.xlsx')

# print tabelaVendas
pd.set_option('display.max_columns', None)
print(tabelaVendas)
print('-' * 50)

# Calcular faturamento por loja
faturamentoLoja = tabelaVendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamentoLoja)
print('-' * 50)

# Calcular quantidade de produtos vendidos por loja
produtosVendidosLoja = tabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(produtosVendidosLoja)
print('-' * 50)

# Calcular ticket médio ( faturamento/quantidade de produtos vendidos)
ticketMedio = (faturamentoLoja['Valor Final'] / produtosVendidosLoja['Quantidade']).to_frame()
ticketMedio = ticketMedio.rename(columns={0: 'Ticket Médio'})
print(ticketMedio)

# envio do e-mail com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'enzocass@live.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue Relatório de Vendas por lojas.</p>

<p>Faturamento:</p>
{faturamentoLoja.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</P>
{produtosVendidosLoja.to_html()}

<p>Ticket Médio:</p>
{ticketMedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Dúvidas estou à disposição.</p>

'''
mail.Send()
