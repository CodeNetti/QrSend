
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
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QInputDialog, QTextEdit, QApplication, QDialog, QPushButton, QVBoxLayout
from PyQt5.QtGui import QIcon
import cv2
import pyautogui
from selenium import webdriver
import time
import numpy as np
import os


def aguardar(caminho_referencia, precisao=0.8, intervalo=1):
    """
    Aguarda até que a imagem especificada seja encontrada na tela.
    
    :param caminho_referencia: Caminho para a imagem de referência.
    :param precisao: Precisão mínima para considerar a imagem encontrada (0.8 por padrão).
    :param intervalo: Intervalo em segundos entre as verificações (1 segundo por padrão).
    :return: True se a imagem for encontrada.
    """
    print("Aguardando a imagem ser localizada na tela...")
    while True:
        # Captura a tela atual
        tela = pyautogui.screenshot()
        tela_np = cv2.cvtColor(np.array(tela), cv2.COLOR_RGB2GRAY)  # Converte para cinza
        imagem_referencia = cv2.imread(caminho_referencia, cv2.IMREAD_GRAYSCALE)  # Lê a imagem de referência em escala de cinza

        # Localiza a imagem na tela
        resultado = cv2.matchTemplate(tela_np, imagem_referencia, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, _ = cv2.minMaxLoc(resultado)

        if max_val >= precisao:
            print("Imagem localizada!")
            return True
        else:
            print("Imagem ainda não encontrada. Tentando novamente...")
            time.sleep(intervalo)


def erro_encontrar(caminho_referencia, precisao=0.8):
    """
    Verifica se a imagem de erro aparece na tela. Caso apareça, clica nela e retorna True.
    
    :param caminho_referencia: Caminho para a imagem de referência.
    :param precisao: Precisão mínima para considerar a imagem encontrada (0.8 por padrão).
    :return: True se o erro foi tratado, False caso contrário.
    """
    print("Verificando erro na tela...")
    
    # Captura a tela atual
    tela = pyautogui.screenshot()
    tela_np = cv2.cvtColor(np.array(tela), cv2.COLOR_RGB2GRAY)  # Converte para cinza
    imagem_referencia = cv2.imread(caminho_referencia, cv2.IMREAD_GRAYSCALE)  # Lê a imagem de referência em escala de cinza

    # Localiza a imagem na tela
    resultado = cv2.matchTemplate(tela_np, imagem_referencia, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(resultado)

    if max_val >= precisao:  # Verifica se encontrou a imagem com a precisão necessária
        x, y = max_loc
        largura, altura = imagem_referencia.shape[::-1]
        centro_x, centro_y = x + largura // 2, y + altura // 2
        pyautogui.moveTo(centro_x, centro_y, duration=0.1)  # Move o cursor para o centro da imagem
        pyautogui.click()  # Realiza o clique
        print("Erro identificado e tratado.")
        return True
    else:
        print("Nenhum erro encontrado.")
        return False
            




# Função para localizar a imagem na tela e retornar as coordenadas
def localizar_imagem_e_clicar(caminho_referencia, precisao=0.8):
    # Captura a tela atual
    tela = pyautogui.screenshot()
    tela_np = cv2.cvtColor(np.array(tela), cv2.COLOR_RGB2GRAY)  # Converte para cinza
    imagem_referencia = cv2.imread(caminho_referencia, cv2.IMREAD_GRAYSCALE)  # Lê a imagem de referência em escala de cinza

    # Localiza a imagem na tela
    resultado = cv2.matchTemplate(tela_np, imagem_referencia, cv2.TM_CCOEFF_NORMED)
    _, max_val, _, max_loc = cv2.minMaxLoc(resultado)

    if max_val >= precisao:  # Verifica se encontrou a imagem com a precisão necessária
        x, y = max_loc
        largura, altura = imagem_referencia.shape[::-1]
        centro_x, centro_y = x + largura // 2, y + altura // 2
        pyautogui.moveTo(centro_x, centro_y, duration=0.3)  # Move o cursor para o centro da imagem
        pyautogui.click()  # Realiza o clique
        return True
    else:
        print("Imagem não encontrada!")
        return False