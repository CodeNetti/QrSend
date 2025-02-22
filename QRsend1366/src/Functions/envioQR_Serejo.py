import qrcode
import json
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage  # Para manipulação da imagem
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib 
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QInputDialog, QTextEdit, QApplication, QDialog, QPushButton, QVBoxLayout
import pyautogui
import requests
import os
import datetime
import pyperclip
from PyQt5.QtGui import QIcon
from Functions.Funcoesdeclique import localizar_imagem_e_clicar, erro_encontrar, aguardar

usuario = os.getlogin()


class InputDialog2(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Digitação")
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Digite o Texto de Envio | <<nomeconvidado>> = Nome do Convidado | <<nomeacomphantes>> = Nome do acompanhantes")
        self.setWindowIcon(QIcon(f'../Pictures/Logo.png'))
        self.submit_button = QPushButton('Pronto')
        self.submit_button.clicked.connect(self.accept)
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addWidget(self.submit_button)
        self.setLayout(layout)
    def get_text(self):
        return self.text_edit.toPlainText()
    
    
def Envio_Original(Dados_Convidados_Envio):

    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setWindowTitle("Texto de envio")
    msg.setText("Por Favor digite o texto que deseja enviar")
    msg.setStandardButtons(QMessageBox.Ok)
    msg.exec_()
    dialog = InputDialog2()
    if dialog.exec_() == QDialog.Accepted:
            texto_digitado = dialog.get_text()

    if texto_digitado:
            # Exibe uma mensagem de confirmação
            caminho_fotoinico = f'../Ver/inicio.png'
            caminho_fotoerro = f'../Ver/erro.png'
            caminho_fotoseta = f'../Ver/seta.png'
            caminho_fotoplus = f'../Ver/plus.png'
            caminho_fotoevidioedoc = f'../Ver/fotovidio.png'
            caminho_fotopesquisa = f'../Ver/pesquisa.png'
            caminho_fotoabrir = f'../Ver/abrir.png'
            caminho_fotoseta2 = f'../Ver/seta2.png'
            caminho_pasta = QFileDialog.getExistingDirectory(None, "Selecione a pasta aonde se encontra os Qrs respectivos da planilha")
            #pyperclip.copy(caminho_envio)
            # Eibe uma mensagem de confirmação
            msg4 = QMessageBox()
            msg4.setIcon(QMessageBox.Information)
            msg4.setWindowTitle("Confirmação")
            msg4.setText("Clique ok para iniciar o disparo")
            msg4.setStandardButtons(QMessageBox.Ok)
            msg4.exec_()
        # Exibe uma mensagem de confirmação
       
    
    listatelefonicadeucerto= []
    listatelefonicadeuerrado= [] 
           
        # Configuração do Selenium
    nav = webdriver.Chrome()
    nav.maximize_window()
    nav.get("https://web.whatsapp.com/")

        # Aguarda o usuário escanear o QR Code manualmente
    aguardar(caminho_fotoinico, precisao=0.8, intervalo=2)

    # Agrupar por número de telefone para enviar todos os QR codes correspondentes em uma única mensagem
    agrupado_telefone = Dados_Convidados_Envio.groupby('Telefone')


    for telefone, grupo in agrupado_telefone:
    # Extrair informações principais do grupo
        nomes = grupo['Nome Convidado'].dropna().unique()  # Remove valores nulos
        nomes = [str(nome) for nome in nomes]  # Converte todos os valores para string
        nome = ', '.join(nomes)  # Converte o array em string separada por vírgulas

        nomes_convidados = grupo['Nome Acompanhante'].dropna().tolist() 
        
        
    # Formatar nomes dos convidados com "e" no último
        if not nomes_convidados:  # Verifica se a lista está vazia
                  nomes_convidados_str =  "" 
        elif len(nomes_convidados) == 1:
                 nomes_convidados_str = nomes_convidados[0]
        elif len(nomes_convidados) == 2:
                 nomes_convidados_str = ' e '.join(nomes_convidados)
        else:
                nomes_convidados_str = ', '.join(nomes_convidados[:-1]) + ' e ' + nomes_convidados[-1]
              
        #nomes_convidados = [str(nome) for nome in nomes_convidados]  # Converte todos os valores para string

        texto_sub = texto_digitado.replace("<<nomeconvidado>>", nome)
        texto_sub = texto_sub.replace("<<nomeacompanhantes>>", nomes_convidados_str)
        textoFormatado = urllib.parse.quote(texto_sub)
        print(telefone)
        listatelefonicadeucerto.append([telefone])             

        #texto = urllib.parse.quote(f"Olá {ConvidadoNome}! Segue o(s) seu(s) QR Code(s).")
        link = f"https://web.whatsapp.com/send?phone={telefone}&text={textoFormatado}"
        nav.get(link)
        time.sleep(10)  
        if erro_encontrar(caminho_fotoerro, precisao=0.8) == True:
            print("Imagem encontrada e clicada!")
            listatelefonicadeuerrado.append([telefone])   

        else:
            localizar_imagem_e_clicar(caminho_fotoseta, 0.8)
            print("Imagem da seta")
            time.sleep(4)
            for _, row in grupo.iterrows():
                print("Imagem da seta deu certo")
                qr_code_path = f'{caminho_pasta}/qrcode_{row["ID"]}.png'
                qr_code_path = f'"{qr_code_path}"'
                qr_code_path = qr_code_path.replace("/", "\\")
                pyperclip.copy(qr_code_path)
                localizar_imagem_e_clicar(caminho_fotoplus, 0.8)
                time.sleep(4)
                localizar_imagem_e_clicar(caminho_fotoevidioedoc, 0.8)
                time.sleep(4)
                localizar_imagem_e_clicar(caminho_fotopesquisa, 0.8)
                time.sleep(1)   
                pyautogui.hotkey("ctrl", "v")
                localizar_imagem_e_clicar(caminho_fotoabrir, 0.8)
                time.sleep(4)
                localizar_imagem_e_clicar(caminho_fotoseta2 , 0.8)
                time.sleep(5)
    wb = Workbook()
    planilha = wb.active  
    planilha.append(["Telefones","Telefones Erros"])
    # Preenche as colunas com as listas
    for i in range(max(len(listatelefonicadeucerto), len(listatelefonicadeuerrado))):
        linha = [
        listatelefonicadeucerto[i][0] if i < len(listatelefonicadeucerto) else "",  # Evita IndexError
        listatelefonicadeuerrado[i][0] if i < len(listatelefonicadeuerrado) else ""
        ]
        planilha.append(linha)
# Salvaa planilha

    data_atual = datetime.datetime.now()

# Formata a data no formato desejado (por exemplo, "AAAA-MM-DD")
    data_formatada = data_atual.strftime("%Y-%m-%d")

# Define o caminho e o nome do arquivo com a data incluída
    caminho_arquivo = f"../Resultados/dadosenvio_{data_formatada}.xlsx"

# Salva o arquivo Excel
    wb.save(caminho_arquivo)
    print("Arquivo Excel criado com sucesso!")



