
import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from PIL import Image as PILImage  # Para manipulação da imagem
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib
from selenium.webdriver.common.by import By
import time
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from PyQt5.QtWidgets import  QMessageBox, QTextEdit, QDialog, QPushButton, QVBoxLayout
from PyQt5.QtGui import QIcon
from selenium import webdriver
import time
import numpy as np
import os
import datetime
from Functions.Funcoesdeclique import localizar_imagem_e_clicar, erro_encontrar, aguardar




usuario = os.getlogin()

class InputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Digitação")
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Digite o Texto de Envio")
        self.setWindowIcon(QIcon(f'C:/Users/{usuario}/Desktop/QRsend/Pictures/Logo.png'))
        self.submit_button = QPushButton('Pronto')
        self.submit_button.clicked.connect(self.accept)
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addWidget(self.submit_button)
        self.setLayout(layout)

    def get_text(self):
        return self.text_edit.toPlainText()
    

class InputDialog2(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Digitação")
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Digite o Texto de Envio | <<nomeconvidado>> = Nome do Convidado | <<nomeacomphantes>> = Nome do acompanhantes")
        self.setWindowIcon(QIcon(f'C:/Users/{usuario}/Desktop/QRsend/Pictures/Logo.png'))
        self.submit_button = QPushButton('Pronto')
        self.submit_button.clicked.connect(self.accept)
        layout = QVBoxLayout()
        layout.addWidget(self.text_edit)
        layout.addWidget(self.submit_button)
        self.setLayout(layout)

    def get_text(self):
        return self.text_edit.toPlainText()

# Função principal
def Envio_Original_Texto(Dados_Convidados_Envio):

    caminho_fotoinico = f'C:/Users/{usuario}/Desktop/QRsend/Ver/inicio.png'
    caminho_fotoerro = f'C:/Users/{usuario}/Desktop/QRsend/Ver/erro.png'
    caminho_fotoseta = f'C:/Users/{usuario}/Desktop/QRsend/Ver/seta.png'

    

    msg = QMessageBox()
    msg.setIcon(QMessageBox.Question)
    msg.setText("Olá, você deseja enviar mensagens nominais?")
    msg.setWindowTitle("Confirmação")
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    resposta = msg.exec_()
    if resposta == QMessageBox.No:
        msg2 = QMessageBox()
        msg2.setIcon(QMessageBox.Information)
        msg2.setWindowTitle("Texto de envio")
        msg2.setText("Por favor digite o texto que deseja enviar")
        msg2.setStandardButtons(QMessageBox.Ok)
        msg2.exec_()
        dialog = InputDialog()
        if dialog.exec_() == QDialog.Accepted:
                texto_digitado = dialog.get_text()

        if texto_digitado:
            # Exibe uma mensagem de confirmação
                msg4 = QMessageBox()
                msg4.setIcon(QMessageBox.Information)
                msg4.setWindowTitle("Confirmação")
                msg4.setText("Clique ok para iniciar o disparo")
                msg4.setStandardButtons(QMessageBox.Ok)
                msg4.exec_()
    
       # Aponto a iamgem, da seta de envio
       
        listatelefonicadeucerto= []
        listatelefonicadeuerrado= []  
  
        # Configuração do Selenium
        nav = webdriver.Chrome()
        nav.maximize_window()
        nav.get("https://web.whatsapp.com/")

        # Aguarda o usuário escanear o QR Code manualmente
        aguardar(caminho_fotoinico, precisao=0.8, intervalo=2)

        agrupado_telefone = Dados_Convidados_Envio.groupby('Telefone')

        for telefone, grupo in agrupado_telefone:
        # Extrair informações principais do grupo
             nomes = grupo['Nome Convidado'].unique()
             nomes_convidados = grupo['Nome Acompanhante'].tolist()

             texto_digitado_refatorado =  urllib.parse.quote(texto_digitado)
             listatelefonicadeucerto.append([telefone])             
             link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto_digitado_refatorado}"    
            # Navegar para o WhatsApp Web
             nav.get(link)
             time.sleep(10)  
             if erro_encontrar(caminho_fotoerro, precisao=0.8) == True:
                    print("Imagem encontrada e clicada!")
                    listatelefonicadeuerrado.append([telefone])             

             else:
            # Usa OpenCV para localizar e clicar na imagem
               localizar_imagem_e_clicar(caminho_fotoseta, 0.8)
               time.sleep(4)
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

# Salva a planilha
        wb.save("dados.xlsx")
        print("Arquivo Excel criado com sucesso!")

    if resposta == QMessageBox.Yes:
        msg2 = QMessageBox()
        msg2.setIcon(QMessageBox.Information)
        msg2.setWindowTitle("Texto de envio")
        msg2.setText("Por favor digite o texto que deseja enviar")
        msg2.setStandardButtons(QMessageBox.Ok)
        msg2.exec_()
        dialog = InputDialog2()
        if dialog.exec_() == QDialog.Accepted:
                texto_digitado = dialog.get_text()

        if texto_digitado:
            # Exibe uma mensagem de confirmaçã

            # Exibe uma mensagem de confirmação
                msg4 = QMessageBox()
                msg4.setIcon(QMessageBox.Information)
                msg4.setWindowTitle("Confirmação")
                msg4.setText("Clique ok para iniciar o disparo")
                msg4.setStandardButtons(QMessageBox.Ok)
                msg4.exec_()
    
    # Aponto a iamgem, da seta de envio
        listatelefonicadeucerto2= []
        listatelefonicadeuerrado2= [] 
           
        # Configuração do Selenium
        nav = webdriver.Chrome()
        nav.maximize_window()
        nav.get("https://web.whatsapp.com/")

        # Aguarda o usuário escanear o QR Code manualmente
        aguardar(caminho_fotoinico, precisao=0.8, intervalo=2)

        agrupado_telefone = Dados_Convidados_Envio.groupby('Telefone')
        

        for telefone, grupo in agrupado_telefone:
    # Extrair informações principais do grupo
             nomes = grupo['Nome Convidado'].dropna().unique()  # Remove valores nulos
             nomes = [str(nome) for nome in nomes]  # Converte todos os valores para string
             nome = ', '.join(nomes)  # Converte o array em string separada por vírgulas
        
             nomes_convidados = grupo['Nome Acompanhante'].dropna().tolist() 
        
             if not nomes_convidados:  # Verifica se a lista está vazia
                  nomes_convidados_str =  "" 
             elif len(nomes_convidados) == 1:
                 nomes_convidados_str = nomes_convidados[0]
             elif len(nomes_convidados) == 2:
                 nomes_convidados_str = ' e '.join(nomes_convidados)
             else:
                 nomes_convidados_str = ', '.join(nomes_convidados[:-1]) + ' e ' + nomes_convidados[-1]
                     
             texto_sub = texto_digitado.replace("<<nomeconvidado>>", nome)
             texto_sub = texto_sub.replace("<<nomeacompanhantes>>", nomes_convidados_str)
             textoFormatado = urllib.parse.quote(texto_sub)

             print(telefone)

             listatelefonicadeucerto2.append([telefone])             
             link = f"https://web.whatsapp.com/send?phone={telefone}&text={textoFormatado}"
             
             nav.get(link)
             time.sleep(10)  
             if erro_encontrar(caminho_fotoerro, precisao=0.8) == True:
                    print("Imagem encontrada e clicada!")
                    listatelefonicadeuerrado2.append([telefone])             

             else:
            # Usa OpenCV para localizar e clicar na imagem
               localizar_imagem_e_clicar(caminho_fotoseta, 0.8)
               time.sleep(4)
        wb = Workbook()
        planilha = wb.active  

        planilha.append(["Telefones","Telefones Erros"])


        # Preenche as colunas com as listas
        for i in range(max(len(listatelefonicadeucerto2), len(listatelefonicadeuerrado2))):
            linha = [
            listatelefonicadeucerto2[i][0] if i < len(listatelefonicadeucerto2) else "",  # Evita IndexError
            listatelefonicadeuerrado2[i][0] if i < len(listatelefonicadeuerrado2) else ""
            ]
            planilha.append(linha)

# Salva a planilha
    data_atual = datetime.datetime.now()

# Formata a data no formato desejado (por exemplo, "AAAA-MM-DD")
    data_formatada = data_atual.strftime("%Y-%m-%d")

# Define o caminho e o nome do arquivo com a data incluída
    caminho_arquivo = f"C:/Users/{usuario}/Desktop/QRsend/Resultados/dadosenvio_{data_formatada}.xlsx"

# Salva o arquivo Excel
    wb.save(caminho_arquivo)
    print("Arquivo Excel criado com sucesso!")