import sys 
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
from Functions.FuncoesInterface import  Gerador_Planilha_Qr_lista_original, Disparo_Qrs_Serejo, Disparo_Texto_Serejo2, Disparo_Texto_Serejo,CriacaodeEventos
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QToolTip, QFileDialog , QMessageBox, QLabel, QDesktopWidget
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel ,QWidget, QVBoxLayout
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5 import QtGui
import os
from PyQt5.QtCore import Qt




usuario = os.getlogin()

class Janela(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Serejo.exe")
        base_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(base_dir)  # Volta uma pasta

        # Define o ícone da janela com caminho relativo
        icon_path = os.path.join(parent_dir, 'Pictures', 'Logoico.ico')
        self.setWindowIcon(QIcon(icon_path))

        # Tamanho da janela
        self.setGeometry(0, 0, 1200, 700)
        self.setStyleSheet("background-color: #222222; color: white;")

        # Centraliza a janela na tela
        self.center()

        # Configura a interface inicial
        self.initUI()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        

    def initUI(self):
        # Adiciona o logo no centro superior
        base_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(base_dir)  # Volta uma pasta
        logo_path = os.path.join(parent_dir, 'Pictures', 'Logo2.png')

        self.logo = QLabel(self)
        pixmap = QPixmap(logo_path)

       
        self.logo.setPixmap(pixmap)
        self.logo.setScaledContents(True)  # Garante que o logo será escalado corretamente
        self.logo.resize(pixmap.width(), pixmap.height())  # Redimensiona para o tamanho real da imagem

            # Centraliza o logo na parte superior da janela
        self.logo.move((self.width() - self.logo.pixmap().width()) // 2, 50)  # Ajuste conforme necessário# Ajuste conforme necessário
        

        # Botões estilizados
        self.addButton1("", 200, 600, self.abrir_janela1)
        self.addButton2("", 600, 600, self.abrir_janela2)

    def addButton(self, text, x, y, func):
        button = QPushButton(text, self)
        button.move(x, y)
        button.resize(120, 100)
        button.setStyleSheet('''
            QPushButton {
                background-color: #ffffff;
                color: #333333;
                font-family: "Roboto";
                font-size: 20px;
                border-radius: 15px;
                padding: 10px;
                box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.2);
            }
            QPushButton:hover {
                background-color: #cccccc;
            }
            QPushButton:pressed {
                background-color: #aaaaaa;
                box-shadow: inset 0px 4px 6px rgba(0, 0, 0, 0.3);
            }
        ''')
        button.clicked.connect(func)
        
    def addButton1(self, text, x, y, func):
        
        label = QLabel("App", self)  # Texto descritivo do botão
        label.move(400,  450)  # Ajuste a posição acima do botão
        label.setStyleSheet('''
            QLabel {
                color: white;
                font-family: "Roboto";
                font-size: 20px;
                text-align: center;
            }
        ''')
        button = QPushButton(text, self)
        button.move(350, 500)
        button.resize(130, 130)
        button.setIcon(QIcon(f"../Pictures/phone.png"))  # Adiciona a imagem como ícone
        button.setIconSize(button.size())
        

         
        
          # Ajusta o ícone ao tamanho do botão  # Adiciona a imagem como ícone
        button.setStyleSheet('''
            QPushButton {
                background-color: #363636;
                color: #333333;
                font-family: "Roboto";
                font-size: 20px;
                border-radius: 15px;
                padding: 0px;
                box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.2);
            }
            QPushButton:hover {
                background-color: #cccccc;
            }
            QPushButton:pressed {
                background-color: #aaaaaa;
                box-shadow: inset 0px 4px 6px rgba(0, 0, 0, 0.3);
            }
        ''')
        button.clicked.connect(func)
        
    def addButton2(self, text, x, y, func):

        label = QLabel("WhatsApp", self)  # Texto descritivo do botão
        label.move(720, 450)  # Ajuste a posição acima do botão
        label.setStyleSheet('''
            QLabel {
                color: white;
                font-family: "Roboto";
                font-size: 20px;
                text-align: center;
            }
        ''')
          # Ajusta manualmente o tamanho do QLabel

        button = QPushButton(text, self)
        button.move(700, 500)
        button.resize(130, 130)
        button.setIcon(QIcon(f"../Pictures/zap.png"))
        button.setIconSize(button.size())  # Ajusta o ícone ao tamanho do botão  # Adiciona a imagem como ícone
        button.setStyleSheet('''
            QPushButton {
                background-color: #363636;
                color: #333333;
                font-family: "Roboto";
                font-size: 20px;
                border-radius: 15px;
                padding: 10px;
                box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.2);
            }
            QPushButton:hover {
                background-color: #cccccc;
            }
            QPushButton:pressed {
                background-color: #aaaaaa;
                box-shadow: inset 0px 4px 6px rgba(0, 0, 0, 0.3);
            }
        ''')
        button.clicked.connect(func)

    def abrir_janela1(self):
        self.janela1 = Interface1()
        self.janela1.show()

    def abrir_janela2(self):
        self.janela2 = Interface2()
        self.janela2.show()


class Interface1(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerar QRs")
        self.setGeometry(0, 0, 1200, 700)
        self.setStyleSheet("background-color: #333333; color: white;")
        self.setWindowIcon(QIcon(f'../Pictures/Logo.png'))

        self.initUI()

    def initUI(self):
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)
        
        # Layout vertical para centralizar os botões
        layout = QVBoxLayout()
        layout.addStretch()  # Espaço flexível acima dos botões
        self.logo = QLabel(self)
        self.logo.setPixmap(QPixmap(f'../Pictures/Logo2.png'))
        self.logo.setScaledContents(True)
        self.logo.resize(200, 200)
        layout.addWidget(self.logo, alignment=Qt.AlignHCenter)  # Centraliza o logo horizontalmente
        layout.addStretch()  # Espaço flexível abaixo dos botões

        self.addButton("Gerar QRs", layout, Gerador_Planilha_Qr_lista_original)
        self.addButton("Criar evento", layout, CriacaodeEventos)
        self.addButton("Realizar envio com os nossos QRs", layout, Disparo_Qrs_Serejo)
        
        layout.addStretch()  # Espaço flexível abaixo dos botões
        self.centralWidget.setLayout(layout)

    def addButton(self, text, layout, func):
        button = QPushButton(text, self)
        button.setFixedSize(320, 100)  # Define um tamanho fixo para manter os botões consistentes
        button.setStyleSheet('''
            QPushButton {
                background-color: #ffffff;
                color: #333333;
                font-family: "Roboto";
                font-size: 20px;
                border-radius: 15px;
                padding: 10px;
                box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.2);
            }
            QPushButton:hover {
                background-color: #cccccc;
            }
            QPushButton:pressed {
                background-color: #aaaaaa;
                box-shadow: inset 0px 4px 6px rgba(0, 0, 0, 0.3);
            }
        ''')
        button.clicked.connect(func)
        layout.addWidget(button, alignment=Qt.AlignHCenter)  # Centraliza horizontalmente o botão


class Interface2(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Envios WhatsApp")
        self.setGeometry(0, 0, 1200, 700)
        self.setStyleSheet("background-color: #333333; color: white;")
        self.setWindowIcon(QIcon(f'../Pictures/Logoico.ico'))

        self.initUI()

    def initUI(self):
        # Widget central e layout vertical
        self.centralWidget = QWidget(self)
        self.setCentralWidget(self.centralWidget)
        layout = QVBoxLayout()
        layout.addStretch()  # Espaço flexível antes do logo

        # Logo centralizado
        self.logo = QLabel(self)
        self.logo.setPixmap(QPixmap(f'../Pictures/Logo2.png'))
        self.logo.setScaledContents(True)
        self.logo.setFixedSize(200, 200)
        layout.addWidget(self.logo, alignment=Qt.AlignHCenter)

        layout.addStretch()  # Espaço flexível entre o logo e os botões

        # Botões centralizados
        self.addButton("Envio de Textos", layout, Disparo_Texto_Serejo)
        self.addButton("Envio de Textos\nDocumentos e Imagens", layout, Disparo_Texto_Serejo2)

        layout.addStretch()  # Espaço flexível após os botões
        self.centralWidget.setLayout(layout)

    def addButton(self, text, layout, func):
        button = QPushButton(text, self)
        button.setFixedSize(320, 100)
        button.setStyleSheet('''
            QPushButton {
                background-color: #ffffff;
                color: #333333;
                font-family: "Roboto";
                font-size: 20px;
                border-radius: 15px;
                padding: 10px;
                box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.2);
                
            }
            QPushButton:hover {
                background-color: #cccccc;
            }
            QPushButton:pressed {
                background-color: #aaaaaa;
                box-shadow: inset 0px 4px 6px rgba(0, 0, 0, 0.3);
            }
        ''')
        button.clicked.connect(func)
        layout.addWidget(button, alignment=Qt.AlignHCenter)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    janela = Janela()
    janela.show()
    sys.exit(app.exec_())
