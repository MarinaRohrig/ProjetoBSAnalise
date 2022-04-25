# Analisar dados e enviar relatório por e-mail.
# Autor: Marina
# Packages: pandas, pywin32 e openpyxl
# Lógica do código:
# Acessar dados de um arquivo XLS, fazer cálculos com os dados do arquivo, gerar um relatório e enviar o relatório via e-mail

# 1. Importar a base de dados, arquivo Vendas.xls;
# 2. Visualizar a base de dados;
# 3. Fazer os cálculos de faturamento da loja. Faturamento = preço de venda x quantidade vendida.
# 4. Fazer os cálculos de quantidade de produto vendidos por loja;
# 5. Fazer os cálculos de ticket médio de cada produto, por loja. O ticket médio é o faturamento / número de vendas;
# 6. Gerar relatórios com base nos resultados dos cálculos;
# 7. Enviar por e-mail os relatórios.

# importa a biblioteca pandas, que lida com planilhas e trata ela como pd
import pandas as pd
# importa a biblioteca que lida com aplicações do windows
import win32com.client as win32

# importa o excel
tabela_vendas = pd.read_excel('Vendas.xlsx')
# mostrar todas as colunas
pd.set_option('display.max_columns', None)
# faturamento
fat = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
# quantidade de prod vend por loja
prodvend = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
# ticket medio e transforma em tabela o resultado com to_frame
ticket_medio = (fat['Valor Final'] / prodvend['Quantidade']).to_frame()
# rename para escrever o nome da coluna que está mostrando valor
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
# testes, pode comentar todos os prints se não quiser os testes
print('*-'*30)
print('Ticket Médio:')
print('*-'*30)
print(ticket_medio)
print('*-'*30)
print('Produtos Vendidos:')
print('*-'*30)
print(prodvend)
print('*-'*30)
print('Faturamento:')
print('*-'*30)
print(fat)
print('*-'*30)
# enviar e-mail, via outlook, precisa ter outlook instalado no pc
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
# marcar para quem vai receber
mail.To = 'marinalarissacarpesrohrig@gmail.com'
# título do e-mail
mail.Subject = 'Relatório de Vendas ProjetoBSAnalise'
# coloquei f na frente para poder usar {} para mostrar variáveis no meio do texto, <p> para criar paragrafos
mail.HTMLBody = f'''
<p> Prezados,</p>
<p> Segue o relatório de vendas, por loja, conforme solicitado. </p>
<p> Faturamento, por loja: </p>
{fat.to_html(formatters={'Valor Final': 'R${:.2f}'.format})}
<p> Quantidade de produtos vendidos, por loja:</p>
{prodvend.to_html()}
<p> Ticket médio dos produtos, em cada loja: </p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:.2f}'.format})}
<p> Para dúvidas estou a disposição. </p>
<p> Att. Marina</p> ~~
'''
mail.Display()
print('E-mail pronto para ser enviado')

