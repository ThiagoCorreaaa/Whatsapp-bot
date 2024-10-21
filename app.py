# Como automatizar esse processo?
# Onde está feito (versão web)
# Tec. preciso pra resolver essa demanda ?
# Payautogui , webbrowser, link whatsapp, oppenpyxl
"""Demanda- Preciso automatizar minhas mensagens p/ meus clientes gostaria de saber valores, e gostaria que entrassem em contato comigo p/
 explicar melhor, quero poder mandar mensagens de cobrança em determinado dia com clientes com vencimento diferente"""

# Passo 1 Descrever os passos manuais e dps transformar isso em código
# Ler planilha e guardar informações sobre nome, telefone e data de vencimento
import openpyxl
from urllib.parse import quote
import webbrowser
import pyautogui
from time import sleep

# Carregar o arquivo do Excel
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # Nome, telefone e vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime("%d/%m/%Y, %H:%M:%S")}. favor pagar no link https//www.link_do_pagamento.com'
    #Link personalizado whatsapp
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'

    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    try:
        seta=pyautogui.locateCenterOnScreen('seta.png')
        sleep(2)
        pyautogui.click(seta[0],seta[1])
        sleep(2)
        pyautogui.hotkey('ctrl','w')
        sleep(2)
    except:
        print (f'Não foi possivel enviar a mensagem para {nome}')
        with open ('erros.csv','a', newline='',encoding='utf-8')as arquivo:
            arquivo.write(f'{telefone},{nome}')