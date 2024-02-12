import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
from PIL import Image

# Abrir imagem da seta
imagem_seta = Image.open('seta.png')

# Ler excel e guardar informações sobre nome, telefone
workbook = openpyxl.load_workbook('clientes.xlsx')
clientes = workbook['Folha1']

for linha in clientes.iter_rows(min_row=2):
    # nome, telefone
    nome = linha[0].value
    telefone = linha[1].value

    mensagem = f'Olá {nome} isto é uma mensagem de teste.'

    try:
        # Criar links personalizados do whatsapp e enviar mensagens para cada contacto
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={
            telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)

        # descobrir a seta para enviar a mensagem
        seta = pyautogui.locateCenterOnScreen(imagem_seta)
        pyautogui.click(seta)

    except:
        # Se não for possivel enviar a mensagem apresenta o contaco no terminal
        # e grava o contacto no ficheiro erros.csv
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')

    # Fechar página para poder abrir uma nova mensagem
    sleep(2)
    pyautogui.hotkey('ctrl', 'w')
