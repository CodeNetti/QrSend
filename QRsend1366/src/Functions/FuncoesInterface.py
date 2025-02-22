
import sys 
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
import time
import urllib.parse
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import pyautogui
import requests
import os
from Functions.geradorQr import  criar_Planilha_Final, insere_Planilha_Final,abrir_Arquivo
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QToolTip, QFileDialog, QMessageBox
from Functions.TextoCru import  Envio_Original_Texto
from Functions.TextoeOutros import  Envio_Original_Texto2
from Functions.CriarEvento   import  criar_Evento  
from Functions.envioQR_Serejo import  Envio_Original


    
def Buscador_carregar_Ler_lista_original():
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Procurar arquivo")
        msg.setText("Escolha a Planilha que deseja gerar os QRs")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
        caminho_arquivo, _ = QFileDialog.getOpenFileName(None, "Selecione o arquivo Excel", "", "Excel Files (*.xlsx *.xls)")
        dados_convidados = pd.read_excel(caminho_arquivo)
        return (dados_convidados)
    

def Gerador_Planilha_Qr_lista_original(self):
    wb, ws = criar_Planilha_Final()
    insere_Planilha_Final(Buscador_carregar_Ler_lista_original(), ws, wb)
    abrir_Arquivo()




def Selecionar_Planilha_QRs_Serejo():
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Procurar arquivo")
        msg.setText("Escolha a Planilha que deseja realizar o disparo")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
        Lista_Disparo, _ = QFileDialog.getOpenFileName(None, "Selecione o arquivo Excel", "", "Excel Files (*.xlsx *.xls)")
        Dados_Convidados_Para_Disparo = pd.read_excel(Lista_Disparo)
        return (Dados_Convidados_Para_Disparo)


#Disparo de textos Crus
def Disparo_Texto_Serejo(self):
        Envio_Original_Texto(Selecionar_Planilha_QRs_Serejo())
#Disparo de textos e outros
def Disparo_Texto_Serejo2(self):
        Envio_Original_Texto2(Selecionar_Planilha_QRs_Serejo())


def Disparo_Qrs_Serejo(self):
       Envio_Original(Selecionar_Planilha_QRs_Serejo())



def Selecionar_Planilha_QRs_QRList():
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Procurar arquivo")
        msg.setText("Escolha a Planilha Gerada Pela QRList")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
        Lista_Disparo, _ = QFileDialog.getOpenFileName(None, "Selecione o arquivo Excel", "", "Excel Files (*.xlsx *.xls)")
        Dados_Convidados_Para_Disparo = pd.read_excel(Lista_Disparo)
        return (Dados_Convidados_Para_Disparo)




       



def CriacaodeEventos(self):
       criar_Evento()
       
       
       
       
    

