from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from numpy import random
import pandas as pd
import PySimpleGUI as sg
import os
import time
import urllib


contatos = pd.read_excel("ana.xlsx")
mensagems = contatos.loc[0, "Mensagem"]

working_directory = os.getcwd()

teste = [  
    [sg.Text("Coloque o caminho da imagem:"),
    sg.InputText(key="-FILE_PATH-"),sg.FileBrowse(initial_folder=working_directory)],
    
]
testeb = [
    [sg.Text("Selecione a opção que deseja:"),
    sg.Combo(['Enviar para grupo sem imagem','Enviar para grupo com imagem','Enviar para numeros sem imagem', 'Enviar para numeros com imagem'],key='dest')],
    
]
testec = [
    [sg.Button('Enviar')]
]
layout = [
     [sg.Column(teste, vertical_alignment='left', element_justification='left')],
     [sg.Column(testeb, vertical_alignment='left', element_justification='left')],
     [sg.Column(testec, vertical_alignment='right', element_justification='right')],
]

window = sg.Window("Web Certificados | Conversor", layout)

def enviar_grupo_mensagem():
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)

    navegador.get("https://web.whatsapp.com/")

    while len(navegador.find_elements(By.ID, 'side')) < 1:
        time.sleep(1)

    for i,message in enumerate(contatos["Grupo"]):
        group = contatos.loc[i, "Grupo"]

        while not(navegador.find_elements(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')):
            time.sleep(1)

        navegador.find_elements(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')
        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').click()
        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(group)
        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(Keys.ENTER)

        while not(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]')):
            time.sleep(1)


        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').click()
        if navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]'):
            time.sleep(1)
            navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').send_keys(mensagems + Keys.ENTER)
            numeroRand = random.randint(10)
            time.sleep(numeroRand)
        else:
            numeroRand = random.randint(10)
            time.sleep(numeroRand)
def enviar_grupo_mensagem_imagem(endereco):
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)

    print(endereco)
    navegador.get("https://web.whatsapp.com/")

    while len(navegador.find_elements(By.ID, 'side')) < 1:
        time.sleep(1)

    for i,message in enumerate(contatos["Grupo"]):
        group = contatos.loc[i, "Grupo"]

        while not(navegador.find_elements(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')):
            time.sleep(1)

        navegador.find_elements(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]')
        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').click()
        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(group)
        navegador.find_element(By.XPATH,'//*[@id="side"]/div[1]/div/div/div[2]/div/div[2]').send_keys(Keys.ENTER)

        while not(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]')):
            time.sleep(1)


        navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').click()
        if navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]'):
            time.sleep(1)
            navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]').send_keys(mensagems + Keys.ENTER)
            midia = endereco
            time.sleep(1)
            navegador.find_element(By.CSS_SELECTOR,"span[data-icon='clip']").click()
            time.sleep(2)
            navegador.find_element(By.CSS_SELECTOR,"input[type='file']").send_keys(midia)
            while not(navegador.find_elements(By.XPATH,'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span')):
                time.sleep(1)
            navegador.find_element(By.XPATH,'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span').click()
            
            numeroRand = random.randint(10)
            time.sleep(numeroRand)
        else:
            numeroRand = random.randint(10)
            time.sleep(numeroRand)

def enviar_numero_mensagem():
    contatos = pd.read_excel("ana.xlsx")
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://web.whatsapp.com/")
 
    while len(navegador.find_elements(By.ID, value='side')) < 1:
        time.sleep(1)

    mensagems = contatos.loc[0, "Mensagem"]

    for i, mensagem in enumerate(contatos["Mensagem"]):
        pessoa = contatos.loc[i, "Pessoa"]
        numero = contatos.loc[i, "Numero"]
        print(numero)

        texto = urllib.parse.quote(f"Olá, {pessoa}\n{mensagems}")
        link = f"https://web.whatsapp.com/send?phone=55{int(numero)}&text={texto}"

        navegador.get(link)
        time.sleep(5)

        if not (navegador.find_elements(By.CLASS_NAME, '_3J6wB')):
            while len(navegador.find_elements(By.ID, value='side')) < 1:
                time.sleep(1)

            while not (navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')):
                time.sleep(1)

            navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()

            numeroRand = random.randint(5,30)
            time.sleep(numeroRand)
        else:
            time.sleep(5)
def enviar_numero_mensagem_imagem(endereco):
    print(endereco)
    contatos = pd.read_excel("ana.xlsx")
    servico = Service(ChromeDriverManager().install())
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://web.whatsapp.com/")
 
    while len(navegador.find_elements(By.ID, value='side')) < 1:
        time.sleep(1)

    mensagems = contatos.loc[0, "Mensagem"]

    for i, mensagem in enumerate(contatos["Mensagem"]):
        pessoa = contatos.loc[i, "Pessoa"]
        numero = int(contatos.loc[i, "Numero"])
        print(numero)
        texto = urllib.parse.quote(f"Olá, {pessoa}\n{mensagems}")
        link = f"https://web.whatsapp.com/send?phone=55{numero}&text={texto}"

        navegador.get(link)
        time.sleep(5)

        if not (navegador.find_elements(By.CLASS_NAME, '_3J6wB')):
            while len(navegador.find_elements(By.ID, value='side')) < 1:
                time.sleep(1)

            while not(navegador.find_elements(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')):
                time.sleep(1)

            navegador.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
            time.sleep(1)
                
            midia = endereco
            navegador.find_element(By.CSS_SELECTOR,"span[data-icon='clip']").click()
            time.sleep(1)
            navegador.find_element(By.CSS_SELECTOR,"input[type='file']").send_keys(midia)

            while not(navegador.find_elements(By.XPATH,'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span')):
                time.sleep(1)

            navegador.find_element(By.XPATH,'//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span').click()

            numeroRand = random.randint(5,30)
            time.sleep(numeroRand)
        else:
            time.sleep(5)


while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Exit'):
        break
    elif event == "Enviar":

        opcao = values["dest"]

        if opcao == 'Enviar para grupo sem imagem':
            enviar_grupo_mensagem()
            
        elif opcao == 'Enviar para grupo com imagem':
            enderecoLen = len(values["-FILE_PATH-"])
            endereco = values["-FILE_PATH-"]

            if enderecoLen > 2:
                enviar_grupo_mensagem_imagem(endereco)
            else:
                sg.popup("Nenhuma imagem foi selecionada")

        elif opcao == 'Enviar para numeros sem imagem':
            enviar_numero_mensagem()

        elif opcao == 'Enviar para numeros com imagem':
            enderecoLen = len(values["-FILE_PATH-"])
            endereco = values["-FILE_PATH-"]
            if enderecoLen > 2:
                enviar_numero_mensagem_imagem(endereco)
            else:
                sg.popup("Nenhuma imagem foi selecionada")
        else:
            sg.popup("Nenhuma opção escolhida")

window.close()