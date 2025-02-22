
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage, ImageDraw, ImageFont
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
import os
from PyQt5.QtWidgets import QFileDialog, QMessageBox,QApplication, QLabel,QVBoxLayout,QDialog, QFileDialog, QMessageBox, QPushButton,QInputDialog
from PyQt5.QtGui import QPixmap , QDesktopServices
from PyQt5.QtCore import QUrl
from openpyxl.drawing.image import Image as ExcelImage
import firebase_admin
from firebase_admin import credentials, firestore
from PyQt5.QtWidgets import QInputDialog, QMessageBox

usuario = os.getlogin()

cred = credentials.Certificate(f'C:/Users/{usuario}/Desktop/QRsend/Pk/serejo-app-firebase-adminsdk-f6b58-095116908a.json')
firebase_admin.initialize_app(cred)

db = firestore.client()

def criar_Evento():
    msgbd = QMessageBox()
    msgbd.setIcon(QMessageBox.Question)
    msgbd.setText("Você deseja criar um evento?")
    msgbd.setWindowTitle("Confirmação")
    msgbd.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    respostaOpen = msgbd.exec_()  # Chama apenas uma vez
    
    if respostaOpen == QMessageBox.Yes:
        EventoTitulo, ok = QInputDialog.getText(None, "Nome do Evento", "Digite o nome do evento:\n Exemplo: Casamento Bruno e Paula")
        
        # Loop para garantir que o nome do evento seja inserido
        while not EventoTitulo and ok:
            msgError = QMessageBox()
            msgError.setIcon(QMessageBox.Warning)
            msgError.setText("Nenhum nome foi inserido, digite o nome do evento.")
            msgError.setWindowTitle("Erro")
            msgError.exec_()
            EventoTitulo, ok = QInputDialog.getText(None, "Nome do Evento", "Digite o nome do evento:\n Ex: Casamento Bruno e Paula")
        
        msg1 = QMessageBox()
        msg1.setIcon(QMessageBox.Information)
        msg1.setWindowTitle("Definir os acessos")
        msg1.setText("Vamos definir o login dos noivos")
        msg1.setStandardButtons(QMessageBox.Ok)
        UserNoivos, okN = QInputDialog.getText(None, "Definição de acesso Noivos", "Defina o login dos noivos no App:\n Exemplo: FESTA33")

        while not UserNoivos and okN:
            msgError = QMessageBox()
            msgError.setIcon(QMessageBox.Warning)
            msgError.setText("Nenhum Login foi inforamdo, defina o login de acesso dos noivos.")
            msgError.setWindowTitle("Erro")
            msgError.exec_()
            UserNoivos, okN = QInputDialog.getText(None, "Definição de acesso Noivos", "Defina o login dos noivos no App:\n Exemplo: FESTA33")

        msg2 = QMessageBox()
        msg2.setIcon(QMessageBox.Information)
        msg2.setWindowTitle("Definir os acessos")
        msg2.setText("Vamos a senha dos noivos")
        msg2.setStandardButtons(QMessageBox.Ok)
        PWNoivos, okW = QInputDialog.getText(None, "Definição de acesso Noivos", "Defina a senha dos noivos no App:\n Exemplo: 123")

        while not PWNoivos and okW:
            msgError = QMessageBox()
            msgError.setIcon(QMessageBox.Warning)
            msgError.setText("Nenhuma senha foi inforamda, defina a senha de acesso dos noivos.")
            msgError.setWindowTitle("Erro")
            msgError.exec_()
            PWNoivos, okW = QInputDialog.getText(None, "Definição de acesso Noivos", "Defina a senha dos noivos no App:\n Exemplo: 123")

        msg3 = QMessageBox()
        msg3.setIcon(QMessageBox.Information)
        msg3.setWindowTitle("Definir os acessos")
        msg3.setText("Vamos definir o login de Usuário")
        msg3.setStandardButtons(QMessageBox.Ok)
        UserUsuarios, okUU = QInputDialog.getText(None, "Definição de acesso de Usuário", "Defina o login de Usuário no App:\n Exemplo: USUARIO13")

        while not UserUsuarios and okUU:
            msgError = QMessageBox()
            msgError.setIcon(QMessageBox.Warning)
            msgError.setText("Nenhum Login foi inforamdo, defina o login de acesso do usuário.")
            msgError.setWindowTitle("Erro")
            msgError.exec_()
            UserUsuarios, okUU = QInputDialog.getText(None, "Definição de acesso de Usuário", "Defina o login de Usuário no App:\n Exemplo: USUARIO13")

        msg4 = QMessageBox()
        msg4.setIcon(QMessageBox.Information)
        msg4.setWindowTitle("Definir os acessos")
        msg4.setText("Vamos definir a senha de Usuário")
        msg4.setStandardButtons(QMessageBox.Ok)
        PWUsuarios, okPWU = QInputDialog.getText(None, "Definição de acesso de Usuário", "Defina a senha de Usuário no App:\n Exemplo: 1234")

        while not PWUsuarios and okPWU:
            msgError = QMessageBox()
            msgError.setIcon(QMessageBox.Warning)
            msgError.setText("Nenhuma senha foi inforamda, defina a senha de acesso do usuário.")
            msgError.setWindowTitle("Erro")
            msgError.exec_()
            PWUsuarios, okPWU = QInputDialog.getText(None, "Definição de acesso de Usuário", "Defina a senha de Usuário no App:\n Exemplo: 1234")

        
        # Se o nome do evento for fornecido, prossegue com a criação
        if ok and okN and okW and okUU and okPWU and EventoTitulo and UserNoivos and PWNoivos and UserUsuarios and PWUsuarios:
            bd_ref = db.collection("EventosTeste")
            
            # Ordenar eventos por nome e pegar o último
            eventoscriados = bd_ref.order_by("idevento", direction=firestore.Query.DESCENDING).limit(1).stream()

            ultimo_evento = None
            for eventoi in eventoscriados:
                ultimo_evento = eventoi
                break

            if ultimo_evento is None:
                novo_numero = 1  # Caso seja o primeiro evento
            else:
                # Extrair o número do nome do último evento (assumindo o formato "EventoXX")
                ultimo_nome = ultimo_evento.id  # Obtemos o ID/documento como "EventoXX"
                numero_evento = int(ultimo_nome.replace("Evento", ""))  # Extraímos o número, removendo o prefixo "Evento"
                novo_numero = numero_evento + 1
            
            novo_evento_nome = f"Evento{novo_numero:02d}"  # Formatar como "EventoXX", ex: "Evento05"
            
            # Referência para o novo documento
            novo_evento_ref = bd_ref.document(novo_evento_nome)
            
            # Dados para o novo evento
            novo_evento_ref.set({
                "EventoTitulo": EventoTitulo,
                "idevento": novo_evento_nome
                # Outros campos podem ser adicionados aqui
            })

            logins_ref = novo_evento_ref.collection("Logins")

            definicao_logins_ref = logins_ref.document("Noivos_Usuarios")

            definicao_logins_ref.set({
             "LoginNoivos": UserNoivos,
             "SenhaNoivos": PWNoivos,
             "LoginUsuarios": UserUsuarios,
             "SenhaUsuarios": PWUsuarios
            })

            msg5 = QMessageBox()
            msg5.setIcon(QMessageBox.Information)
            msg5.setWindowTitle("Inserção de Convidados")
            msg5.setText("Selecione a planilha de convidados")
            msg5.setStandardButtons(QMessageBox.Ok)
            msg5.exec_()

    # Abre o explorador de arquivos para o usuário selecionar a planilha
            Arquivo, _ = QFileDialog.getOpenFileName(None, "Selecione o arquivo Excel", "", "Excel Files (*.xlsx *.xls)")

    # Se o arquivo foi selecionado, abre o arquivo com o aplicativo padrão do sistema
            if Arquivo:
                QDesktopServices.openUrl(QUrl.fromLocalFile(Arquivo))

                # Exibe a mensagem de confirmação para fazer o upload da planilha
                msgopen = QMessageBox()
                msgopen.setIcon(QMessageBox.Question)
                msgopen.setText("Deseja fazer o upload da planilha selecionada?")
                msgopen.setWindowTitle("Confirmação")
                msgopen.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
                rspopen = msgopen.exec_()

                # Enquanto a resposta for "No", permite que o usuário selecione outra planilha
                while rspopen == QMessageBox.No:
                    msg6 = QMessageBox()
                    msg6.setIcon(QMessageBox.Information)
                    msg6.setWindowTitle("Inserção de Convidados")
                    msg6.setText("Selecione a planilha de convidados")
                    msg6.setStandardButtons(QMessageBox.Ok)
                    msg6.exec_()
                    Arquivo, _ = QFileDialog.getOpenFileName(None, "Selecione outro arquivo Excel", "", "Excel Files (*.xlsx *.xls)")
                    if Arquivo:
                        QDesktopServices.openUrl(QUrl.fromLocalFile(Arquivo))

                        # Pergunta novamente se deseja fazer o upload da nova planilha selecionada
                        msgopen.setText("Deseja fazer o upload dos dados dessa planilha selecionada?")
                        rspopen = msgopen.exec_()

                # Se o usuário confirmar com "Yes", continue com o processo de upload
                if rspopen == QMessageBox.Yes:
                    df = pd.read_excel(Arquivo)

                    


                    count_sucesso = 0  # Contador de inserções bem-sucedidas
                    count_falha = 0  
                
                # Iterar sobre as linhas da planilha
                    df = df.where(pd.notna(df), None)

                    count_sucesso = 0  # Contador de inserções bem-sucedidas
                    count_falha = 0 
                    
                    
                    df = pd.read_excel(Arquivo, dtype={"Idade": str})  # Carregar coluna Idade como string
                    df["Idade"] = df["Idade"].astype(str)  # Garantir que todos os valores sejam string
 # Contador de falhas na inserção
                    df = df.where(pd.notna(df), None)  # Substituir valores NaN por None (padrão Python)

                    # Iterar sobre as linhas do DataFrame
                    for idx, row in df.iterrows():
                        convidados_ref = novo_evento_ref.collection("Convidados")
                        convidado_id = f"convidado{idx+1:02d}"

                        try:
                            # Inserir cada convidado na sub-coleção "Convidados"
                            convidados_ref.document(convidado_id).set({
                             "ID": row.iloc[0],  # Posição 0
                             "Nome Acompanhante": row.iloc[2] if len(row) > 2 else None,  # Verifique se há pelo menos 3 colunas
                             "Telefone": row.iloc[4] if len(row) > 4 else None,  # Verifique se há pelo menos 4 colunas
                             "Nome Convidado": row.iloc[1] if len(row) > 1 else None,  # Verifique se há pelo menos 2 colunas
                             "Nome QrCode": row.iloc[3] if len(row) > 3 else None,  # Verifique se há pelo menos 2 colunas
                             "Faixa Etaria": row.iloc[6] if len(row) > 6 else None,  # Verifique se há pelo menos 6 colunas
                            "Idade": row.iloc[7] if row.iloc[7] else None,  # Garantir que seja None se vazio
                             "Javalidado": False,
                              # Coluna 2: Telefone         
                               # Coluna 4
                        })
                            count_sucesso += 1
                        except Exception as e:
                            print(f"Erro ao inserir convidado {convidado_id}: {e}")
                            count_falha += 1

                    # Exibir contadores de sucesso e falhas
                    msg_resultado = QMessageBox()
                    msg_resultado.setIcon(QMessageBox.Information)
                    msg_resultado.setWindowTitle("Resultado do Upload")
                    msg_resultado.setText(f"Upload concluído!\n\nConvidados inseridos com sucesso: {count_sucesso}\nConvidados com falha: {count_falha}")
                    msg_resultado.exec_()

                             # Incrementar falhas
                   

                    # Aqui, você pode chamar a função que fará o upload ou processar a planilha.







            
            print(f"Novo evento criado: {novo_evento_nome} com o título '{EventoTitulo}'")
        else:
            print("Operação cancelada ou Dado Faltando.")
