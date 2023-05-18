# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# 01 => IMPORTACOES DE MODULOS

import sys
import csv
import shutil
import pandas as pd
import numpy as np
import os
import re
import PyQt5.QtWidgets as qtw
import PyQt5.QtCore as qtc
from PyQt5 import uic

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# 02 => INTERFACE GRAFICA - GUI (Graphical User Interface)

class UI(qtw.QMainWindow):

    def __init__(self):
        super(UI, self).__init__()

        # Carrega o arquivo GUI
        root_folder = os.path.dirname(os.path.abspath(__file__))
        GUI_file = os.path.join(root_folder, "GUI_Interface.ui")
        uic.loadUi(GUI_file, self)

        # Adiciona um nome a janela
        # self.setWindowTitle("PCP Usinagem - Programa Auxiliar")

        # Define os widgets 
        self.tab_Gerador = self.findChild(qtw.QTabWidget, "tab_Gerador")
        self.pushButton_ProcurarPCP = self.findChild(qtw.QPushButton, "pushButton_ProcurarPCP")
        self.pushButton_ProcurarDesenhos = self.findChild(qtw.QPushButton, "pushButton_ProcurarDesenhos")
        self.pushButton_GerarPastas = self.findChild(qtw.QPushButton, "pushButton_GerarPastas")
        self.pushButton_Atualizar = self.findChild(qtw.QPushButton, "pushButton_Atualizar")
        self.pushButton_LimparTudo = self.findChild(qtw.QPushButton, "pushButton_LimparTudo")
        self.pushButton_reserva = self.findChild(qtw.QPushButton, "pushButton_reserva")
        self.radioButton_codigo = self.findChild(qtw.QRadioButton,"radioButton_codigo")
        self.radioButton_referencia = self.findChild(qtw.QRadioButton, "radioButton_referencia")
        self.checkBox_Pintura = self.findChild(qtw.QCheckBox, "checkBox_Pintura")
        self.checkBox_Usinagem = self.findChild(qtw.QCheckBox, "checkBox_Usinagem")
        self.checkBox_Processos = self.findChild(qtw.QCheckBox, "checkBox_Processos")
        self.lineEdit_ArquivoPCP = self.findChild(qtw.QLineEdit, "lineEdit_ArquivoPCP")
        self.lineEdit_Desenhos = self.findChild(qtw.QLineEdit, "lineEdit_Desenhos")
        self.tab_ImportarLista = self.findChild(qtw.QTabWidget, "tab_ImportarLista")
        self.pushButton_ArquivoPMS = self.findChild(qtw.QPushButton, "pushButton_ArquivoPMS")
        self.pushButton_ArquivoPCPPadrao = self.findChild(qtw.QPushButton, "pushButton_ArquivoPCPPadrao")
        self.pushButton_GerarPCP = self.findChild(qtw.QPushButton, "pushButton_GerarPCP")
        self.pushButton_VerificarArquivos = self.findChild(qtw.QPushButton, "pushButton_VerificarArquivos")
        self.lineEdit_ArquivoPMS = self.findChild(qtw.QLineEdit, "lineEdit_ArquivoPMS")
        self.lineEdit_ArquivoPCPPadrao = self.findChild(qtw.QLineEdit, "lineEdit_ArquivoPCPPadrao")
        self.lineEdit_NomeProjeto = self.findChild(qtw.QLineEdit, "lineEdit_NomeProjeto")
        self.lineEdit_NomeLista = self.findChild(qtw.QLineEdit, "lineEdit_NomeLista")
        self.StatusPanel = self.findChild(qtw.QPlainTextEdit, "StatusPanel")

        # Acoes dos widgets
        self.pushButton_ProcurarPCP.clicked.connect(self.ProcurarPCP)
        self.pushButton_ProcurarDesenhos.clicked.connect(self.ProcurarDesenhos)
        self.pushButton_GerarPastas.clicked.connect(self.GerarPastas)
        self.pushButton_Atualizar.clicked.connect(self.Atualizar)
        self.pushButton_LimparTudo.clicked.connect(self.LimparTudo)
        self.pushButton_reserva.clicked.connect(self.GerarPlanilhas)
        self.pushButton_ArquivoPCPPadrao.clicked.connect(self.ArquivoPCPPadrao)
        self.pushButton_GerarPCP.clicked.connect(self.GerarPCP)
        self.pushButton_VerificarArquivos.clicked.connect(self.VerificarArquivos)

        self.lineEdit_ArquivoPCP.textChanged.connect(self.liberar_botoes1)
        self.lineEdit_Desenhos.textChanged.connect(self.liberar_botoes1)
        self.radioButton_codigo.toggled.connect(self.liberar_botoes1)
        self.radioButton_referencia.toggled.connect(self.liberar_botoes1)
        self.lineEdit_ArquivoPMS.textChanged.connect(self.liberar_botoes2)
        self.lineEdit_ArquivoPCPPadrao.textChanged.connect(self.liberar_botoes2)
        self.lineEdit_NomeProjeto.textChanged.connect(self.liberar_botoes2)
        self.lineEdit_NomeLista.textChanged.connect(self.liberar_botoes2)

        self.clean_button_cleaner = bool(False)

        # Redireciona a saída da função Print para o widget StatusPanel
        sys.stdout.write = self.StatusPanel.insertPlainText

        # +++++++++++++++++++++++++++++++++++
        # + DECLARACAO DE VARIAVEIS GLOBAIS +
        # +++++++++++++++++++++++++++++++++++

        # Cria a instânica de configurações gerais para uso nas funções
        self.settings = qtc.QSettings()

        # Cria o argumento inicial "Ultima pasta acessada" contendo o caminho do arquivo python
        self.settings.setValue("Ultima Pasta Acessada", os.path.dirname(os.path.abspath(__file__)))

        # Cria variavel para manter ultima pasta acessada
        # self.ultima_pasta_acessada = ()

        self.PCP_path = ()
        self.ListaPCP = ()


        # Apresenta a janela do aplicativo
        self.show()

# +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# 04 => FUNCOES

    # Bloqueia/desbloqueia os botoes de acao da tab1
    def liberar_botoes1(self):

        LinhaTextoPCP = self.lineEdit_ArquivoPCP.text()
        LinhaTextoDesenhos = self.lineEdit_Desenhos.text()
        CB_Processos = True if self.checkBox_Processos.isChecked() else False
        CB_Pintura = True if  self.checkBox_Pintura.isChecked() else False
        CB_Usinagem = True if  self.checkBox_Usinagem.isChecked() else False

        if LinhaTextoPCP:
            self.pushButton_LimparTudo.setEnabled(True)
        else:
            self.pushButton_LimparTudo.setEnabled(False)

        if LinhaTextoPCP and LinhaTextoDesenhos:
            if CB_Processos or CB_Pintura or CB_Usinagem:
                self.pushButton_GerarPastas.setEnabled(True)
                self.pushButton_Atualizar.setEnabled(True)
            else:
                self.pushButton_GerarPastas.setEnabled(False)
                self.pushButton_Atualizar.setEnabled(False)

    # Bloqueia/desbloqueia os botoes de acao da tab2
    def liberar_botoes2(self):
        leC = self.lineEdit_ArquivoPMS.text()
        leD = self.lineEdit_ArquivoPCPPadrao.text()
        leE = self.lineEdit_NomeProjeto.text()
        leF = self.lineEdit_NomeLista.text()

        if leC and leD and leE and leF:
            self.pushButton_GerarPCP.setEnabled(True)
        else:
            self.pushButton_GerarPCP.setEnabled(False)

        if leE and leF:
            self.pushButton_VerificarArquivos.setEnabled(True)
        else:
            self.pushButton_VerificarArquivos.setEnabled(False)

    # Abre a janela para selecionar o arquivo de PCP
    def ProcurarPCP(self):
        # Abre a janela de interação com o usario para procurar o arquivo de PCP
        fname1 = qtw.QFileDialog.getOpenFileName(
            parent=self,
            caption="Selecione um arquivo",
            directory=self.settings.value("Ultima Pasta Acessada"),
            filter="Excel Files (*.xlsx *.xlsm)"
        )
        # Escreve o caminho do arquivo selecionado na linha correspondente da interface
        if fname1[0]:
            self.lineEdit_ArquivoPCP.setText(str(fname1[0]))
            print("Caminho do arquivo PCP selecionado:", fname1[0])
            ultima_pasta_acessada = os.path.dirname(fname1[0])
            self.settings.setValue("Ultima Pasta Acessada", ultima_pasta_acessada)
        else:
            print("Nenhum arquivo foi selecionado.")
            return None
        return fname1

    # Abre a janela para selecionar a pasta contendo os arquivos das peças
    def ProcurarDesenhos(self):
        # Open file dialog
        fname2 = qtw.QFileDialog.getExistingDirectory(
            parent=self,
            caption="Selecione o local dos desenhos",
            directory=self.settings.value("Ultima Pasta Acessada"),
            options=qtw.QFileDialog.ShowDirsOnly
        )
        # Escreve o caminho na linha
        if fname2:
            fname2 = qtc.QDir.cleanPath(fname2)
            self.lineEdit_Desenhos.setText(fname2)
            print("Caminho dos desenhos: ", fname2)
        else:
            print("Nenhuma pasta foi selecionada.")
            return None
        return fname2

    # Função que cria as pastas
    def GerarPastas(self):

        # Carrega o arquivo Excel e seleciona as colunas desejadas
        self.PCP_Path = (self.lineEdit_ArquivoPCP.text())
        # noinspection PyTypeChecker
        self.ListaPCP = pd.read_excel(
            io=self.PCP_path,
            sheet_name="Principal",
            usecols="A,C,G,J,K"
        )

        # Cria as listas com os arquivos de desenhos existente
        self.Desenhos_Path = (self.lineEdit_Desenhos.text())
        self.ListaArquivos = os.listdir(self.Desenhos_Path)

        # -------------------------------------------------------------------------------

        if self.checkBox_Processos.isChecked():
            self.pastaprocessos()
        else:
            print("Pastas de PROCESSOS não foram criadas, pois não foi solicitado.")

        # Cria a pasta de pintura completa caso a opcao seja selecionada
        if self.checkBox_Pintura.isChecked():
            self.pastapintura()
        else:
            print("Pasta de PINTURA não foi criada, pois não foi solicitado.")

        # Cria a pasta de usinagem completa caso a opcao seja selecionada
        if self.checkBox_Usinagem.isChecked():
            self.pastausinagem()
        else:
            print("Pasta de USINAGEM não foi criada, pois não foi solicitado.")

    def pastaprocessos(self):

        # Cria a lista de processos sem repetição, sem desmembrar processos compostos
        lista_processos = list(set(self.ListaPCP['G'].dropna().tolist()))

        # Separa os processos compostos e cria a lista final
        lista_processos_temporaria = []  # Lista vazia para armazenar os itens separados
        for item in lista_processos:
            if '+' in item:
                substrings = item.split("+")
                lista_processos_temporaria.extend(substrings)
            else:
                lista_processos_temporaria.append(item)
        lista_processos = list(set(lista_processos_temporaria))

        # Mostra resultado da lista final
        print_processos = ', '.join(str(x) for x in lista_processos)
        print("Lista de processos:", print_processos)

        if self.radioButton_codigo.isChecked():
            planilhaPCP = self.ListaPCP.drop(labels=['C', 'J'], axis=1)
        elif self.radioButton_referencia.isChecked():
            planilhaPCP = self.ListaPCP.drop(labels=['A', 'J'], axis=1)

        else:
            print("Coluna com os códigos dos desenhos não foi selecionada. Favor selecionar uma opção válida.")
            return

        # Itera por cada processo, criando a pasta e copiando os desenhos correpondentes
        for processo in lista_processos:

            pasta_processo = os.path.join((os.path.dirname(self.PCP_Path)), processo)  # Nome da pasta a ser criada
            # Confere se a pasta do processo já existe antes de criá-la. Caso ela já exista, o programa informa o usuário e pula para a próxima iteração
            if os.path.exists(pasta_processo):
                print("A pasta "{}' já existe, ela não será criada novamente e este processo não será criado. Utilize o botão Atualizar caso queira sobrescrever os arquivos e receber um relatório dos arquivos alterados.'.format(processo))
                continue
            else:
                os.mkdir(pasta_processo)  # Cria a pasta do processo

            # Filtra a planilha pelo processo correspondente da iteracao
            codigos_por_processo = planilhaPCP.loc[planilhaPCP['G'].str.contains(processo, case=False, na=False)]
            codigos = []  # Cria uma variavel vazia tipo lista para conter os codigos do processo

            # Orienta o programa a buscar os códigos de acordo com a informaçao do usuario
            if self.radioButton_codigo.isChecked():
                codigos = list(codigos_por_processo['A'])
            elif self.radioButton_referencia.isChecked():
                codigos = list(codigos_por_processo['C'])

            # Procura os arquivos correspondentes a cada código e copia para a pasta
            for codigo in codigos:
                # Garante que a varíavel 'código' é do tipo string
                codigo = str(codigo)
                # Cria uma lista com os arquivos que contém o código em seu nome
                arquivos = [arquivo for arquivo in self.ListaArquivos if re.search(codigo, arquivo, re.I)]

                #Copia cada um dos arquivos para dentro da pasta
                for arquivo in arquivos:
                    arquivo_futuro = os.path.join(pasta_processo, arquivo)  # Cria o nome do arquivo futuro

                    # Confere se o desenho a ser copiado ja existe
                    if os.path.exists(arquivo_futuro):
                        print('O arquivo "{}" já existe na pasta e não será copiado.'.format(arquivo))
                        continue
                    else:
                        arquivo_original = os.path.join(self.Desenhos_Path, arquivo)
                        shutil.copy(arquivo_original, pasta_processo)
        print("Pastas de PROCESSOS criadas com sucesso")

    def pastapintura(self):

        # Filtra a planilha para conter apenas as linhas com itens de pintura
        listapintura = self.ListaPCP.loc[self.ListaPCP['J'].str.contains("FP|RAL|pintar", case=False, na=False)]

        # Cria lista com os códigos das peças de pintura, dependendo da seleção do usuário
        if self.radioButton_codigo.isChecked():
            listapintura = list(set(listapintura['A'].tolist()))
        elif self.radioButton_referencia.isChecked():
            listapintura = list(set(listapintura['C'].tolist()))


        # Caminho para pasta de pintura
        pasta_pintura = os.path.join((os.path.dirname(self.PCP_Path)), "PINTURA")
        # Verifica se a pasta já existe
        if os.path.exists(pasta_pintura):
            print('A pasta "PINTURA" já existe, o procedimento será abortado.')
            return
        else:
            os.mkdir(pasta_pintura)  # Cria a pasta de pintura

        arquivos_pintura = []  # Variavel lista com os arquivos de desenhos de pintura encontrados

        # Procura os arquivos com base nos codigos
        for codigo in listapintura:
            codigo = str(codigo)
            arquivos = [arquivo for arquivo in self.ListaArquivos if re.search(codigo, arquivo, re.I)]
            arquivos_pintura.extend(arquivos)

        # Copia os arquivos encontrados
        for arquivo in arquivos_pintura:
            arquivo_pintura_final = os.path.join(pasta_pintura, arquivo)  # Cria o nome do arquivo futuro
            if os.path.exists(arquivo_pintura_final):  # Confere se o desenho a ser copiado ja existe
                print('O arquivo "{}" já existe na pasta e não será copiado.'.format(arquivo))
            else:
                arquivo_pintura_original = os.path.join(self.Desenhos_Path, arquivo)
                shutil.copy(arquivo_pintura_original, pasta_pintura)
        print('Pasta de "PINTURA" criada com sucesso.')

    def pastausinagem(self):

        # Cria a lista de processos sem repetição, sem desmembrar processos compostos
        lista_usinagem = set(self.ListaPCP['K'].dropna().tolist())
        lista_usinagem.discard("-")
        lista_usinagem = list(lista_usinagem)

        # Caminho da pasta de Usinagem
        pasta_usinagem = os.path.join((os.path.dirname(self.PCP_Path)), "USINAGEM")
        # Verifica se a pasta já existe
        if os.path.exists(pasta_usinagem):
            print('A pasta "USINAGEM" já existe, o procedimento será abortado.')
        else:
            os.mkdir(pasta_usinagem)  # Cria a pasta de usinagem

        if self.radioButton_codigo.isChecked():
            planilhaPCP = self.ListaPCP.drop(labels=['C', 'K'], axis=1)
        elif self.radioButton_referencia.isChecked():
            planilhaPCP = self.ListaPCP.drop(labels=['A', 'K'], axis=1)

        for porte_usinagem in lista_usinagem:
            # Nome da pasta a ser criada
            pasta_porte_usinagem = os.path.join(pasta_usinagem, porte_usinagem)  # Nome da pasta a ser criada
            # Verifica se a pasta já existe
            if os.path.exists(pasta_porte_usinagem):
                print("A pasta "{}' já existe, ela não será criada novamente e este processo não será criado. Utilize o botão Atualizar caso queira sobrescrever os arquivos e receber um relatório dos arquivos alterados.'.format(porte_usinagem))
                break
            else:
                os.mkdir(pasta_porte_usinagem)  # Cria a pasta do processo

            # Filtra a planilha pelo processo correspondente da iteracao
            codigos_por_usinagem = planilhaPCP.loc[planilhaPCP['K'].str.fullmatch(porte_usinagem, case=False, na=False)]
            codigos_usinagem = []  # Cria uma variavel vazia tipo lista para conter os codigos do processo

            if self.radioButton_codigo.isChecked():
                codigos_usinagem = list(codigos_por_usinagem['A'])
            elif self.radioButton_referencia.isChecked():
                codigos_usinagem = list(codigos_por_usinagem['C'])

            for codigo_usinagem in codigos_usinagem:
                codigo_usinagem = str(codigo_usinagem)
                arquivos = [arquivo for arquivo in self.ListaArquivos if re.search(codigo_usinagem, arquivo, re.I)]  # Separa os arquivos iguais aos codigos
                for arquivo in arquivos:
                    arquivo_futuro_usinagem = os.path.join(pasta_porte_usinagem, arquivo)  # Cria o nome do arquivo futuro
                    if os.path.exists(arquivo_futuro_usinagem):  # Confere se o desenho a ser copiado ja existe
                        break
                    else:
                        arquivo_origem_usinagem = os.path.join(self.Desenhos_Path, arquivo)
                        shutil.copy(arquivo_origem_usinagem, pasta_porte_usinagem)
        print("Pasta de USINAGEM criada com sucesso.")

    def Atualizar(self):
        pass

    def LimparTudo(self):
        if self.clean_button_cleaner == False:
            self.pushButton_LimparTudo.setText("Confirma?")
            self.clean_button_cleaner = True
        elif self.clean_button_cleaner == True:
            xlsmflfldr = (self.lineEdit_ArquivoPCP.text())
            filesdir = os.path.dirname(xlsmflfldr)
            foldersdir = pd.DataFrame(os.walk(filesdir))
            foldersdir = (foldersdir[0].squeeze())
            if foldersdir:
                for item in foldersdir:
                    if not item == filesdir:
                        shutil.rmtree(item, ignore_errors=True)
                print("Todas as pastas foram excluídas com sucesso!")
                self.clean_button_cleaner = False
                self.pushButton_LimparTudo.setText("Limpar Tudo")

    def GerarPlanilhas(self):
        pass

    def ArquivoPMS(self):
        # Open file dialog
        pmspath = qtw.QFileDialog.getOpenFileName(
            parent=self,
            caption="Abrir Arquivo",
            directory=os.path.dirname(os.path.abspath(__file__)),
            filter="CSV Files (*.csv)"
        )

        # Output filename to line.edit
        if pmspath:
            self.lineEdit_ArquivoPMS.setText(str(fname1[0]))

            # Output diretory to status panel
            print("Caminho do arquivo PMS: ", (fname1[0]))

        return pmspath

    def ArquivoPCPPadrao(self):
        # Open file dialog
        PCPPadraopath = qtw.QFileDialog.getOpenFileName(
            parent=self,
            caption="Abrir Arquivo",
            directory=os.path.dirname(os.path.abspath(__file__)),
            filter="XLSM Files (*.xlsm)"
        )

        # Output filename to line.edit
        if PCPPadraopath:
            self.lineEdit_ArquivoPCPPadrao.setText(str(fname1[0]))

            # Output diretory to status panel
            print("Caminho do arquivo PCP Padrão: ", (fname1[0]))

        return PCPPadraopath

    def GerarPCPTESTE(self):
        pmspath = (self.lineEdit_ArquivoPMS.text())
        pmsfile = pd.read_csv(
            filepath_or_buffer=pmspath,
            header=0,
            usecols="B,C,D,G,H,I,M,N,O",
        )
        pmsfile = pmsfile.loc[pmsfile['I'].str.contains('BR-PROD-MA')]

    def GerarPCP(self):
        pmspath = (self.lineEdit_ArquivoPMS.text())
        PCPPadraopath = (self.lineEdit_ArquivoPCPPadrao.text())



    def VerificarArquivos(self):
        pass

app = qtw.QApplication(sys.argv)
UIWindow = UI()

# Run the app
app.exec_()
