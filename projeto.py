# Importar a base de dados
import pandas as pd
tabelaVendas = pd.read_excel("Vendas.xlsx")


# Visualizar a base
pd.set_option('display.max_columns', None)
print(tabelaVendas)


# Faturamento por loja
faturamento = tabelaVendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)


# Quantidade de produtos vendidos por loja
# Pegar a ID Loja e Quantidade
# Agrupar por ID Loja e somar as quantidades

quantidade_produtos_vendidos = tabelaVendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_produtos_vendidos)
print('-' * 50)


# Ticket médio por produto em cada loja
# Isso é o faturamento / quantidade_produtos_vendidos
ticket_medio = (faturamento['Valor Final'] / quantidade_produtos_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns = {0: 'Ticket Médio'})
print(ticket_medio)


# Enviar um e-mail pelo relatório
import win32com.client as win32

outlook = win32.Dispatch('outlook.application') #conectando o python e o outlook
mail = outlook.CreateItem(0) #cria um e-mail
mail.To = 'endereçoDeEmail' # para quem vai enviar
mail.Subject = 'Message subject'
mail.HTMLBody = f'''
<p>
Prezado(a)s,<br>
Segue o Relatório de Vendas por cada loja.
</p>

<h4>Faturamento:</h4>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<h4>Quantidade vendida:</h4>
{quantidade_produtos_vendidos.to_html()}

<h4>Ticket Médio do produtos em cada loja:</h4>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou a disposição.</p><br>
<p>Att., nome.</p>
f'''

mail.Send()
print('E-mail enviado.')