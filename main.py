"""
automatizar leitura de dados da planilha - openpyxl
teclado - pyautogui
acesso ao site (simples) - webbrowser
digitação automatizada - link whatsapp

"""

import openpyxl
import openpyxl
import pyautogui as pg
from urllib.parse import quote
import webbrowser
from time import sleep
# abrir e logar no whatsapp
webbrowser.open("https://web.whatsapp.com/")
sleep(15)

# acessar a planilha
workbook = openpyxl.load_workbook("data/contatos.xlsx")
sheet_contatos = workbook["Contatos"]

for linha in sheet_contatos.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    mensagem = linha[2].value
    try:
        link_msg = f"https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}"

        webbrowser.open(link_msg)
        sleep(10)

        enviar_mensagem = pg.locateCenterOnScreen("data/seta_whatsapp.png")
        sleep(5)

        pg.click(enviar_mensagem[0], enviar_mensagem[1], duration=1)
        sleep(5)

        pg.hotkey("ctrl", "w")
        sleep(2)
        pg.hotkey("ctrl", "w")
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}.')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}')
