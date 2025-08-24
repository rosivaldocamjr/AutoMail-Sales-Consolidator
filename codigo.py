# leia o arquivo informacoes.txt para entender o 
# projeto e os passos iniciais

# os: usado p/ controlar e automatizar tarefas no sistema de 
# arquivos e no ambiente do computador
import os

# pandas: usado p/ trabalhar com grandes volumes de dados
import pandas as pd

# pywin32: enviar email
import win32com.client as win32

# biblioteca datetime do python importando a ferramenta datetime 
from datetime import datetime

# caminho da pasta arquivos.csv
# exemplo: C:\Users\rosivaldo\Downloads
caminho = "arquivos_csv"

# pegar todos os arquivos da pasta
# os.listdir() é a forma de ver o que existe dentro de uma pasta
arquivos = os.listdir(caminho)

# consolidar (juntar) os arquivos em 1 só
# criar tabela vazia com pd.DataFrame()
# DataFrame é uma tabela do python, aqui ela está vazia
tabela_consolidada = pd.DataFrame()

# a variavel arquivos é uma lista com os nomes dos arquivos
# os.path.join() = concatena os caminhos
# pd.read_csv() = lê os arquivos csv
for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    
    # concatena a tabela_vendas a tabela_consolidada
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

# ordenar a tabela pela data de venda
tabela_consolidada = tabela_consolidada.sort_values(by="data de venda")

# quando executamos o programa, vemos um indice a esquerda no terminal
# eu quero resetar esses indices para aparecer do jeito certo
# e os indices antigos iremos excluir com o drop=True
tabela_consolidada = tabela_consolidada.reset_index(drop=True)

# exportando a tabela_consolidada para excel
# retirando o indice com index=False
tabela_consolidada.to_excel("Vendas.xlsx", index=False)

# conecta o python com o outlook
# o computador chama o outlook de outlook.application
# para esse comando funcionar o outlook deve está configurado e funcionando no pc
outlook = win32.Dispatch('outlook.application')

# cria um email que você pode preencher e enviar
email = outlook.CreateItem(0)

# pra quem enviar o email
email.To = "rosivaldo.bannerjr@gmail.com"

data_hoje = datetime.today().strftime("%d/%m/%Y")

# assunto do email
email.Subject = f"Relatório de Vendas {data_hoje}"

# corpo do email
email.Body = f"""
Prezados,

Segue em anexo o Relatório de Vendas de {data_hoje} atualizado.
Qualquer dúvida estou a disposição.
Abs,
Rosivaldo
"""

# pega o caminho onde o Python está rodando e guarda na variável caminho
caminho = os.getcwd()

# junta o caminho da pasta atual com o nome do arquivo 
# Vendas.xlsx para formar o caminho completo do arquivo
anexo = os.path.join(caminho, "Vendas.xlsx")

# adicionar o anexo ao email
email.Attachments.Add(anexo)

# enviar email
email.Send()