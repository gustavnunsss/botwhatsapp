import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(30) # Espera de 30 segundos
# Ler planinha e guardar informações sobre nome, telefone e data de vencimento 

workbook = open.load_workbook('Lista_Contatos_Pagamentos.pdf')
paginas_clientes = workbook = [''] # Nome da pág que está a planilha

for linha in paginas_clientes.iter_rows(min_row=2):
  # nome, telefone, vencimento  
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y, %H:%M:%S')}. Favor pagar no link'
    # https//www.link_do_pagamento.com

# Criar links personalizados do whatsapp e enviar mensagens para cada cliente 
# Com base nos dados da planilha 
try:
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}' webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    seta = pyautogui.locateOnScreen('Screenshot 2025-05-06 23.39.27.png')
    sleep(5)
    pyautogui.click(seta[0],seta[1]) # Será ultilizado para clicar na imagem
    sleep(5)
    pyautogui.hotkey('ctrl','w')
    sleep(5)
except:
    print(f'Não foi possivel enviar mensagem para {nome}') # Para quando não der certo, ele guarada dados das mensagens que não foram enviadas
    with open('erros.csv', 'a',newline='',encoding='utf-8')
        arquivo.write(f'{nome}, {telefone}')