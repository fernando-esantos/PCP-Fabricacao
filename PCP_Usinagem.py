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

        # Le o arquivo e cria a lista de processos, sem desmembrar processos compostos
        lista_processos = list(set(self.ListaPCP['PROCESSOS'].dropna().tolist()))

        # Separa os processos compostos e cria lista final
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

        ###### Cria as listas de desenhos existentes
        ###### desenhos_path = (self.lineEdit_Desenhos.text())
        ###### arqdir = os.listdir(desenhos_path)

        if self.radioButton_codigo.isChecked():
            planilhaPCP = self.ListaPCP.drop(labels=['C', 'J'], axis=1)

        elif self.radioButton_referencia.isChecked():
            planilhaPCP = self.ListaPCP.drop(labels=['A', 'J'], axis=1)

        else:
            print("Coluna com os códigos dos desenhos nao foi selecionada. Favor informar uma opcao valida.")
            return

        tipos_arquivos = [".igs", ".pdf", ".dwg"]

        # Cria as pastas (directories)
        for processo in lista_processos:
            folder_path = os.path.join((os.path.dirname(self.PCP_Path)), processo)  # Nome da pasta a ser criada
            if os.path.exists(folder_path):
                print("A pasta ->", processo, "<- já existe, ela não será criada novamente e este processo não será "
                                              "criado. Utilize o botão Atualizar caso queira sobrescrever os arquivos"
                                              " e receber um relatório dos arquivos alterados")
                continue
            else:
                os.mkdir(folder_path)  # Cria a pasta do processo

            # Separa a coluna de codigos pelo processo a iteracao
            codigos_por_processo = planilhaPCP.loc[planilhaPCP['PROCESSOS'].str.contains(processo)]
            codigos = []  # Cria uma variavel vazia tipo lista para conter os codigos do processo

            # Orienta o programa a buscar os códigos de acordo com a informaçao do usuario
            if self.radioButton_codigo.isChecked():
                codigos = list(codigos_por_processo['CODIGO'])
            elif self.radioButton_referencia.isChecked():
                codigos = list(codigos_por_processo['REFERENCIA'])

            for codigo in codigos:
                codigo = str(codigo)
                arquivos = [arquivo for arquivo in self.ListaArquivos if re.search(codigo, X, re.I)]  # Separa os arquivos iguais aos codigos
                arquivos_encontrados = list(arquivos)  # Cria uma lista com os arquivos encontrados

                for arquivo_encontrado in arquivos_encontrados:
                    nome_arquivo = os.path.splitext(arquivo_encontrado)[0]
                    extensao_arquivo = os.path.splitext(arquivo_encontrado)[1]
                    if extensao_arquivo in tipos_arquivos and codigo in nome_arquivo:
                        # Adicionar o arquivo correspondente à lista
                        caminho_completo_origem = os.path.join(root, file)
                        arquivos_correspondentes.append(caminho_completo_origem)
                    dest = os.path.join(folder_path, arquivo_encontrado)  # Cria o nome do arquivo futuro

                    # Confere se o desenho a ser copiado ja existe
                    if os.path.exists(dest):
                        print("O arquivo {} já existe na pasta e não será copiado", arquivo_encontrado)
                        continue
                    else:
                        origem = os.path.join(desenhos_path, arquivo_encontrado)
                        shutil.copy(origem, folder_path)
        print("Pastas de PROCESSOS criadas com sucesso")

    def pastapintura(self):

        ltacbmnt = pd.read_excel(PCP_path, sheet_name="Principal", usecols="J")
        ltacbmnt = ltacbmnt.loc[ltacbmnt['ACAB. SUPERFICIAL'].str.contains('FP', na=False)]
        ltacbmnt = np.array(ltacbmnt)
        ltacbmnt = (np.unique(ltacbmnt))
        ltpntr = ()
        if self.radioButton_codigo.isChecked():
            ltpntr = pd.read_excel(PCP_path, sheet_name="Principal", usecols="A,J", skipfooter=1)
        elif self.radioButton_referencia.isChecked():
            ltpntr = pd.read_excel(PCP_path, sheet_name="Principal", usecols="C,J", skipfooter=1)
        cdpntr = [] # Objeto lista com os codigos dos desenhos de pintura
        flpntr = []  # Objeto lista com os arquivos de desenhos de pintura
        fdpntr = os.path.join((os.path.dirname(PCP_path)), "PINTURA") # Nome para pasta de pintura
        if os.path.exists(fdpntr):
            print("A pasta ->PINTURA<- já existe, o procedimento será abortado.")
            return
        else:
            os.mkdir(fdpntr)  # Cria a pasta de pintura
        for item in ltacbmnt:
            pntr_temp = ltpntr.loc[ltpntr['ACAB. SUPERFICIAL'].str.contains(item, na=False)]
            lcdpntr = []
            if self.radioButton_codigo.isChecked():
                lcdpntr = list(pntr_temp['CODIGO'])
            elif self.radioButton_referencia.isChecked():
                lcdpntr = list(pntr_temp['REFERENCIA'])
            cdpntr.extend(lcdpntr)
        for cd in cdpntr:
            cd = str(cd)
            srchflpntr = [fl for fl in arqdir if re.search(cd, fl, re.I)]
            tltpntr = list(srchflpntr)
            flpntr.extend(tltpntr)
        for files in flpntr:
            ftfile = os.path.join(fdpntr, files)  # Cria o nome do arquivo futuro
            if os.path.exists(ftfile):  # Confere se o desenho a ser copiado ja existe
                print(files + "já existe.")
            else:
                rgnfile = os.path.join(desenhos_path, files)
                shutil.copy(rgnfile, ftfile)
        print("Pasta de PINTURA criada com sucesso.")

    def pastausinagem(self):

        # Listagem e separacao dos processos de usinagem
        ltsngm = pd.read_excel(PCP_path, sheet_name="Principal", usecols="K")
        ltsngm = ltsngm.replace("-", "@")
        ltsngm = ltsngm[ltsngm != '@']
        ltsngm = ltsngm.dropna()
        ltsngm = ltsngm.squeeze()
        ltsngm = ltsngm.unique()

        # Cria a pasta de Usinagem
        sngmfdpth = os.path.join((os.path.dirname(PCP_path)), "USINAGEM")
        if os.path.exists(sngmfdpth):
            self.StatusPanel.append("A pasta ->USINAGEM<- já existe, o procedimento será abortado.")
        else:
            os.mkdir(sngmfdpth)  # Cria a pasta de usinagem
        for pieces in ltsngm:
            sngnmfldr = pieces.replace(" / ", "+")
            sngprcssfd = os.path.join(sngmfdpth, sngnmfldr)  # Nome da pasta a ser criada
            # self.StatusPanel.append(sngprcssfd)
            if os.path.exists(sngprcssfd):
                self.StatusPanel.append(str("A pasta ->" + sngnmfldr + "<- já existe,, o procedimento será abortado."))
                break
            else:
                os.mkdir(sngprcssfd)  # Cria a pasta do processo
            sngprcsssplt = lc.loc[lc['PORTE USINAGEM'].str.fullmatch(pieces, na=False)]  # Separa a coluna de codigos por processo
            sngcds = []
            if self.radioButton_codigo.isChecked():
                sngcds = list(sngprcsssplt['CODIGO'])
            elif self.radioButton_referencia.isChecked():
                sngcds = list(sngprcsssplt['REFERENCIA'])
            for YY in sngcds:
                YY = str(YY)
                sngfls = [XX for XX in arqdir if re.search(YY, XX, re.I)]  # Separa os arquivos iguais aos codigos
                sngfls = list(sngfls)  # Cria uma lista com os arquivos encontrados
                for WW in sngfls:
                    sngflsdstntn = os.path.join(sngprcssfd, WW)  # Cria o nome do arquivo futuro
                    if os.path.exists(sngflsdstntn):  # Confere se o desenho a ser copiado ja existe
                        break
                    else:
                        sngflsrgn = os.path.join(desenhos_path, WW)
                        shutil.copy(sngflsrgn, sngflsdstntn)
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
