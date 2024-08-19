#passo a passo -> código

import openpyxl
from urllib.parse import quote
import webbrowser 
from time import sleep
import pyautogui 


webbrowser.open('https://web.whatsapp.com/') 
sleep(30)



#Ler planilha e guardar informações sobre  nome, telefone e data de vencimento 
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    mensagem = f'Olá {nome} o boleto esta vencendo,{vencimento.strftime('%d%m%Y')}. Pague na chave pix (11)98208-0831'

#Criar Links personalizados do whatsApp e enivar mensagens para cada cliente 
#Com base nos dados da planilha
    link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    webbrowser.open(link_mensagem_whatsapp)
    sleep(10)
    try:
        seta = pyautogui.locateCenterOnScreen('seta2.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except: 
        print(f'Nao foi possivel enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')