# AutoExcel
Cadastro de Vendas e Automação de Preenchimento Este repositório contém dois scripts em Python e dois arquivos Excel para gerenciamento de vendas:

Cadastro: Este script permite cadastrar informações de vendas em uma planilha do Excel por meio de uma interface gráfica simples usando Tkinter. As informações incluem cliente, produto, quantidade e categoria do produto.

Automação: Este script automatiza o preenchimento de dados em outra planilha Excel. Ele lê dados de uma planilha e utiliza a biblioteca PyAutoGUI para preencher automaticamente informações como cliente, produto, quantidade e categoria em uma aplicação externa.

Como Usar:

Certifique-se de ter as dependências Python instaladas:
• tkinter
• openpyxl
• pyautogui (apenas para o script de automação)

Execute o script cadastro.py para abrir a interface de cadastro e preencher os dados na planilha Vazio.xlsx.
Para usar o script de automação, certifique-se de ter as coordenadas corretas para os cliques do mouse (verifique as coordenadas no script auto.py). Execute o script auto.py para automatizar o preenchimento dos dados na planilha vendas_de_produtos.xlsx.

Arquivos:

• cadastro.py: Script Python para cadastro de vendas.
• auto.py: Script Python para automação do preenchimento de dados.
• Vazio.xlsx: Planilha Excel vazia para o cadastro de vendas.
• vendas_de_produtos.xlsx: Planilha Excel utilizada para automação do preenchimento de dados.

Notas Importantes:

• Certifique-se de que as dependências estão instaladas corretamente.
• Para o script de automação, é importante que o mouse esteja nas coordenadas corretas durante a execução.
• Os arquivos Excel são utilizados para armazenar e gerenciar os dados de vendas.
