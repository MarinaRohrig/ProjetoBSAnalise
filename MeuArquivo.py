# Analisar dados e enviar relatório por e-mail.
# Autor: Marina

# Lógica do código:
# Acessar dados de um arquivo XLS, fazer cálculos com os dados do arquivo, gerar um relatório e enviar o relatório via e-mail

# 1. Importar a base de dados, arquivo Vendas.xls;
# 2. Visualizar a base de dados;
# 3. Fazer os cálculos de faturamento da loja. Faturamento = preço de venda x quantidade vendida.
# 4. Fazer os cálculos de quantidade de produto vendidos por loja;
# 5. Fazer os cálculos de ticket médio de cada produto, por loja. O ticket médio é o faturamento / número de vendas;
# 6. Gerar relatórios com base nos resultados dos cálculos;
# 7. Enviar por e-mail os relatórios.

import pandas as pd
#vai importar a biblioteca pandas, que lida com planilhas e trata ela como pd
