import openpyxl  # Biblioteca para trabalhar com arquivos do Excel
import pyautogui as pg  # Biblioteca para automação de tarefas no computador
from urllib.parse import quote  # Função para codificar URLs
import webbrowser  # Biblioteca para abrir links no navegador
from time import sleep  # Função para pausar a execução por um tempo determinado

# Abrir e logar no WhatsApp Web
webbrowser.open("https://web.whatsapp.com/")
sleep(15)  # Pausa para permitir o tempo de login no WhatsApp Web

# Acessar a planilha de contatos
workbook = openpyxl.load_workbook("data/contatos.xlsx")
sheet_contatos = workbook["Contatos"]

# Iterar sobre as linhas da planilha, começando da segunda linha (min_row=2)
for linha in sheet_contatos.iter_rows(min_row=2):
    nome = linha[0].value  # Captura o nome na primeira coluna
    telefone = linha[1].value  # Captura o telefone na segunda coluna
    mensagem = linha[2].value  # Captura a mensagem na terceira coluna
    
    try:
        # Criar o link para enviar a mensagem via WhatsApp Web
        link_msg = f"https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}"
        
        # Abrir o link no navegador
        webbrowser.open(link_msg)
        sleep(10)  # Esperar a página carregar

        # Localizar o botão de envio da mensagem no WhatsApp Web
        enviar_mensagem = pg.locateCenterOnScreen("data/seta_whatsapp.png")
        sleep(5)  # Pausa antes de clicar no botão

        # Clicar no botão para enviar a mensagem
        pg.click(enviar_mensagem[0], enviar_mensagem[1], duration=1)
        sleep(5)  # Pausa para garantir que a mensagem foi enviada

        # Fechar a aba do navegador (duas vezes para garantir o fechamento)
        pg.hotkey("ctrl", "w")
        sleep(2)
        pg.hotkey("ctrl", "w")
        sleep(5)  # Pausa antes de processar a próxima linha

    except:
        # Caso ocorra um erro, imprimir uma mensagem de erro
        print(f'Não foi possível enviar mensagem para {nome}.')
        
        # Registrar o erro em um arquivo CSV
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}\n')
