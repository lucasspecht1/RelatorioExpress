#importa as bibliotecas para funcionamento do programa
import pandas as pd
import files
import os
import gn

from threading import Thread
from PyQt5 import uic, QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from openpyxl import Workbook
from openpyxl.styles import Font, Color, Alignment
from datetime import date


# Verifica o tema do programa para poder inicializar
arquivo_tema1 = open('data.txt','r')
arquivo_tema = arquivo_tema1.read()
arquivo_tema1.close()

if arquivo_tema == 'c':
    app=QtWidgets.QApplication([])
    tela=uic.loadUi('untitled.ui')

elif arquivo_tema == 'e':
    app=QtWidgets.QApplication([])
    tela=uic.loadUi('untitled_b.ui')

else:
    arquivo_tema = open('data.txt','w')
    arquivo_tema.write('c')
    arquivo_tema.close
    app=QtWidgets.QApplication([])
    tela=uic.loadUi('untitled.ui')
   
# Define os parametros de interface para layout
sucesso = uic.loadUi('dialog_sucess.ui')
sucesso.show()
sucesso.hide()


warning = uic.loadUi('dialog_warning.ui')
warning.show()
warning.hide()


error = uic.loadUi('dialog_error.ui')
error.show()
error.hide()


tela.setFixedSize(960,720)
tela.setWindowIcon(QtGui.QIcon('icon1.png'))

#-> Inicia o gif de carregamento
move = QMovie('loading1.gif')
tela.label_6.setMovie(move)
tela.label_6.setScaledContents(1)
move.start()

#-> Enconde elementos não necessarios atualmente
tela.label_2.hide()
tela.label_3.hide()
tela.label_6.hide()
tela.dateEdit.hide()
tela.dateEdit_2.hide()

#-> Define elementos (combo Box)
tela.comboBox.addItems(["","Intervalo Personalizado","24 Horas","7 Dias","31 Dias"])


# Define a função de tema claro
def claro():
    arquivo_tema = open('data.txt','w')
    arquivo_tema.write('c')
    arquivo_tema.close
    tela.close()
    os.startfile('main.pyw')
    
#-> Define a função de tema Escuro
def escuro():
    arquivo_tema = open('data.txt','w')
    arquivo_tema.write('e')
    arquivo_tema.close
    tela.close()
    os.startfile('main.pyw')

def loading(tela):
    tela.label_6.show()
    tela.toolButton.hide()
    
#Define função de pesquisa principal e suas variaveis de uso
def busca_principal():
    
    #-> Argumentos planilha
    wb = Workbook()
    planilha = wb.worksheets[0]

    fonte = Font(name='Calibri',size=12,bold=True)  
    fonte_h = Font(underline='single')
    
    planilha['A1'] = "Titulo:"
    planilha['B1'] = "Veiculo:"
    planilha['C1'] = "Data:"
    planilha['D1'] = "Links:"
    planilha['A1'].font = fonte
    planilha['B1'].font = fonte
    planilha['C1'].font = fonte
    planilha['D1'].font = fonte

    planilha.column_dimensions['A'].width = 50
    planilha.column_dimensions['C'].width = 16
    planilha.column_dimensions['A'].height = 45
    planilha.column_dimensions['B'].width = 24
    planilha.column_dimensions['D'].width = 10

    letra_celula = "A"
    letra_celula2 = "B"
    letra_celula3 = "C"
    letra_celula4 = "D"

    num_celula = "2"
    num_celula2 = "2"
    num_celula3 = "2"
    num_celula4 = "2"


    celula = letra_celula + num_celula
    celula2 = letra_celula2 + num_celula2
    celula3 = letra_celula3 + num_celula3
    celula4 = letra_celula4 + num_celula4

    c_g = 0
    contador = 0
    contador_d = 0
    n = 2

    status = True

    #-> Variaveis de condições de busca
    data = tela.dateEdit.date()
    data2 = tela.dateEdit_2.date()
    
    data_final = str(data.toPyDate()).replace('-','/')
    data_final2 = str(data2.toPyDate()).replace('-','/')
    
    cliente = str(tela.lineEdit.text())
    periodo = int(tela.comboBox.currentIndex())
    page = 0
    
        
    #-> Local para exportação (caso não tenha salvo ele cria)
    try:
        arquivo = open('localfile.txt','r')
        arquivo = arquivo.read()

    except:
        file = open('localfile.txt','w')
        local = QtWidgets.QFileDialog.getSaveFileName()[0]
        localf = local+".xlsx"
        file.write(localf)
        file.close()
        arquivo = str(localf)
        
    #-> Condicionais para a busca
    if periodo == 1:

        resultado = gn.buscar(f'{cliente}+after:{data_final}+before:{data_final2}')
        df = pd.DataFrame(resultado)

        try:
            title = df['titulo']

            link = df['link']
        
            data = df['data']
            
            site = df['veiculo']
            
        except: 
            print('Nenhum Resultado')
            warning.show()
            #--> Retorna um status de busca vazio
            status = False
            #--> Popup de erro para busca sem retorno
            #criar popup
            tela.label_6.hide()
            tela.toolButton.show()
    
    elif periodo == 2:
        
        resultado = gn.buscar(f'{cliente}+when:1d')
        df = pd.DataFrame(resultado)
        
        try:
            title = df['titulo']

            link = df['link']
        
            data = df['data']

            site = df['veiculo']
            
        except: 
            print('Nenhum Resultado')
            warning.show()
            #--> Retorna um status de busca vazio
            status = False
            #--> Popup de erro para busca sem retorno
            #criar popup
            tela.label_6.hide()
            tela.toolButton.show()
                
    elif periodo == 3:
        
        resultado = gn.buscar(f'{cliente}+when:7d')
        df = pd.DataFrame(resultado)
        
        try:
            title = df['titulo']

            link = df['link']
        
            data = df['data']

            site = df['veiculo']

        except:
            print('Nenhum Resultado')
            warning.show()
            #--> Retorna um status de busca vazio
            status = False
            #--> Popup de erro para busca sem retorno
            #criar popup
            tela.label_6.hide()
            tela.toolButton.show()
            
    elif periodo == 4:

        resultado = gn.buscar(f'{cliente}+when:31d')
        df = pd.DataFrame(resultado)
        
        try:
            title = df['titulo']

            link = df['link']
        
            data = df['data']

            site = df['veiculo']

        except:
            print('Nenhum Resultado')
            warning.show()
            #--> Retorna um status de busca vazio
            status = False
            #--> Popup de erro para busca sem retorno
            #criar popup
            tela.label_6.hide()
            tela.toolButton.show()
            
    while True:
    #-> Loop para criação de cada celula da planilha com suas informações    
        try:    
            planilha[celula] = title[contador]
            planilha[celula].alignment =  Alignment(wrap_text=True)
            planilha[celula2] = site[contador]
            planilha[celula2].alignment =  Alignment(wrap_text=True)
            planilha[celula3] = data[contador]
            planilha[celula3].alignment =  Alignment(wrap_text=True)
            planilha[celula4].hyperlink = f"{link[contador]}"
            planilha[celula4].value = 'Notícia'
            planilha[celula4].font = fonte_h        
            planilha[celula4].alignment =  Alignment(wrap_text=True)
            
            num_celula = int(num_celula)
            num_celula2 = int(num_celula2)
            num_celula3 = int(num_celula3)
            num_celula4 = int(num_celula4)
            
            num_celula4 += 1
            num_celula3 += 1
            num_celula2 += 1
            num_celula += 1
            
            num_celula = str(num_celula)
            num_celula2 = str(num_celula2)
            num_celula3 = str(num_celula3)
            num_celula4 = str(num_celula4)
            
            celula = letra_celula + num_celula
            celula2 = letra_celula2 + num_celula2
            celula3 = letra_celula3 + num_celula3
            celula4 = letra_celula4 + num_celula4

            contador += 1
            c_g += 1
        except:
            print('e')
            break
        
    #-> Salva a planilha e finaliza com popup
    if status == True:
        print('ok2')
        try:
            planilha.title = "Relatorio express"
            wb.save(arquivo)
            
            tela.label_6.hide()
            tela.toolButton.show()
            sucesso.show()
            print('salvo')

        except:
            error.show()
            tela.label_6.hide()
            tela.toolButton.show()
            print('erro ao salvar arquivo!')
                 

        
#Faz a configuração e geração de um popup com informações do programa
def popup_sobre():

    popup = QMessageBox()
    popup.setWindowTitle('Sobre')
    popup.setText('Programa de Versão 0.3       ⠀\n\n By: Deeper Technology')
    popup.setIcon(QMessageBox.Information)
    popup.setWindowIcon(QtGui.QIcon('icon1.png'))
    popup.exec_()


#Cria um loop para a escolha de intervalo e o aparecimento de calendarios
def selecao_intervalo():

    while True:
        if tela.comboBox.currentText() == "Intervalo Personalizado":
            tela.label_2.show()
            tela.label_3.show()
            tela.dateEdit.show()
            tela.dateEdit_2.show()
            
        else:
            tela.label_2.hide()
            tela.label_3.hide()
            tela.dateEdit.hide()
            tela.dateEdit_2.hide()

                
#Cria as condicionais para busca
def botao_busca():
    #-> Cria um popup para avisos
    popup = QMessageBox()
    popup.setWindowTitle('Atenção')
    popup.setIcon(QMessageBox.Warning)
    popup.setWindowIcon(QtGui.QIcon('icon1.png'))

    #->Encontra cliente e data de pesquisa
    cliente = tela.lineEdit.text()
    opcao_intervalo = tela.comboBox.currentIndex()

    if cliente == "":
        popup.setText('Por Favor digite um cliente valido! ⠀')
        popup.exec_()

    elif opcao_intervalo == -1:
        popup.setText('Por favor selecione um periodo de pesquisa! ⠀')
        popup.exec_()

    else:
        tela.toolButton.hide()
        tela.label_6.show()
        t2 = Thread(target=busca_principal)
        t2.start()
        
        #time.sleep(1)
        #busca_principal()
        
        
#Encontra o local para salvar o arquivo
def local_arquivo():
    file = open('localfile.txt','w')
    local = QtWidgets.QFileDialog.getSaveFileName()[0]
    localf = local+".xlsx"
    file = open('localfile.txt','w')
    file.write(localf)
    file.close()

        
#Inicia a interface e thread(funções em paralelismo)

#-> Configurações de tela
tela.actionOp_es.triggered.connect(local_arquivo)
tela.actionSobre.triggered.connect(popup_sobre)
tela.toolButton.clicked.connect(botao_busca)
tela.actionClaro.triggered.connect(claro)
tela.actionEscuro.triggered.connect(escuro)
#tela.actionEscuro.triggered.connect(tema_escuro)


t = Thread(target=selecao_intervalo)

tela.show()
t.start()
app.exec()
       

