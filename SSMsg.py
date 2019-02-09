#! /usr/bin/python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import base64
import json
import traceback
import platform

from zeep import Client  # doit être installé en plus de python à l'aide pip

from platform import system as system_name  # Returns the system/OS name
from subprocess import call as system_call  # Execute a shell

from PyQt5 import QtGui
from PyQt5 import QtCore
from PyQt5 import QtWidgets


class Gui(QtWidgets.QMainWindow):

    def __init__(self):
        super(Gui, self).__init__()
        self.dictConfig = {}
        self.dictEleves = {}
        self.current_dir = os.path.expanduser("~")

        ######################################
        #paramétrage de la fenêtre principale#
        ######################################
        #self.setGeometry(300, 300, 300, 200)
        # le titre de la fenêtre principale
        self.setWindowTitle('SmartSchool Messages')
        # centre la fenêtre
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        # fin du centrage
        # attribue une icone à la fenêtre principale
        self.setWindowIcon(QtGui.QIcon('./icons/ssmsg.svg'))

        ######################
        #Création des actions#
        ######################

        # importFilesAction : ouvre une fenêtre de dialogue pour importer les fichiers à envoyer
        importFilesAction = QtWidgets.QAction(QtGui.QIcon(
            './icons/import_file.svg'), "&Sélectionner les fichiers des à envoyer", self)
        importFilesAction.setShortcut('Ctrl+O')
        importFilesAction.setStatusTip("Sélectionner les fichiers")
        importFilesAction.triggered.connect(self.verifAvantImport)

        # quitAction : fermeture du logiciel
        quitAction = QtWidgets.QAction(
            QtGui.QIcon('./icons/quit.svg'), "Quitter", self)
        quitAction.setShortcut('Ctrl+Q')
        quitAction.setStatusTip("Quitter")
        quitAction.triggered.connect(self.appExit)

        # aboutAction : ouvre une fenêtre d'information à propos de ce logiciel, qui l'a écrit, avecc quel langage, à quoi il sert
        aboutAction = QtWidgets.QAction(QtGui.QIcon(
            './icons/ssmsg.svg'), "Á propos", self)
        aboutAction.setStatusTip("À propos de ce logiciel")
        aboutAction.triggered.connect(self.about)

        # helpAction : ouvre une fenêtre d'information sur la personne à contacter en cas de problème
        helpAction = QtWidgets.QAction(QtGui.QIcon(
            './icons/help.svg'), "Obtenir de l'aide", self)
        helpAction.setStatusTip(
            "Obtenir de l'aide concernant l'utilisation de ce logiciel")
        helpAction.triggered.connect(self.aide)

        # configAction : ouvre une fenêtre de dialogue pour configurer le logiciel
        configAction = QtWidgets.QAction(QtGui.QIcon(
            './icons/config.svg'), "Configuration", self)
        configAction.setStatusTip("Configurer ce logiciel")
        configAction.triggered.connect(self.openDialConfig)

        # refreshDictAction : recharge le dictEleves depuis SS
        refreshAction = QtWidgets.QAction(QtGui.QIcon(
            './icons/refresh.svg'), "Recharger liste élèves", self)
        refreshAction.setStatusTip(
            "Recharger le dictionnaire depuis SmartSchool")
        refreshAction.triggered.connect(self.refreshDictEleves)

        # création d'une barre d'outil qui sera placée dans la fenetre principale
        self.toolbar = self.addToolBar('Toolbar')
        # ajout des actions dans la barre d'outils
        self.toolbar.addAction(importFilesAction)
        self.toolbar.addAction(quitAction)
        self.toolbar.addAction(configAction)
        self.toolbar.addAction(aboutAction)
        self.toolbar.addAction(helpAction)
        self.toolbar.addAction(refreshAction)

        #################
        #le Dock Message#
        #################
        # création du Dock
        dockWidgetMessage = QtWidgets.QDockWidget(self)
        dockWidgetMessage.setTitleBarWidget(QtWidgets.QLabel(
            '<p style="font-size:11pt;font-weight:bold">Message</p>'))
        dockWidgetMessage.setFeatures(
            QtWidgets.QDockWidget.NoDockWidgetFeatures)

        # création d'un layout vertical et d'un widget dans lequel on place le layout
        layout = QtWidgets.QVBoxLayout()
        widget = QtWidgets.QWidget()
        widget.setLayout(layout)

        # création des widget du message
        self.sujet = QtWidgets.QLineEdit()
        self.message = QtWidgets.QTextEdit()
        self.boutton_ok = QtWidgets.QPushButton(
            QtGui.QIcon('./icons/ok_apply.svg'), 'Envoyer')
        self.boutton_ok.setMinimumHeight(40)
        self.boutton_ok.clicked.connect(self.verifAvantEnvoi)

        # création des checkbox pour l'envoi des messages
        self.sendComptePrincipal = QtWidgets.QCheckBox("Compte Principal")
        self.sendComptePrincipal.setToolTip(
            '<p>Envoyer le fichier sur le compte principal</p>')
        self.sendComptePrincipal.setCheckState(0)
        self.sendComptePrincipal.stateChanged.connect(self.changeDestinataire)
        self.sendComptesSecondaires = QtWidgets.QCheckBox(
            "Comptes secondaires")
        self.sendComptesSecondaires.setToolTip(
            ('<p>Envoyer le fichier sur les comptes secondaires 1 et 2</p>'))
        self.sendComptesSecondaires.setCheckState(2)
        self.sendComptesSecondaires.stateChanged.connect(
            self.changeDestinataire)

        # placement des widget du message dans le layout
        layout.addWidget(QtWidgets.QLabel("<p><b>Sujet</b></p>"))
        layout.addWidget(self.sujet)
        layout.addWidget(QtWidgets.QLabel("<p><b>Message</b></p>"))
        layout.addWidget(self.message)
        layout.addWidget(QtWidgets.QLabel("<p><b>Envoi</b></p>"))
        layout.addWidget(self.sendComptePrincipal)
        layout.addWidget(self.sendComptesSecondaires)
        layout.addStretch(1)
        layout.addWidget(self.boutton_ok)

        # placement du widget dans le dock
        dockWidgetMessage.setWidget(widget)

        # placement du dock dans la fenêtre, à gauche
        self.addDockWidget(QtCore.Qt.LeftDockWidgetArea, dockWidgetMessage)
    #

    def ping(host):
        """
        Returns True if host (str) responds to a ping request.
        Remember that a host may not respond to a ping (ICMP) request even if the host name is valid.
        """

        # Ping command count option as function of OS
        param = '-n' if system_name().lower() == 'windows' else '-c'

        # Building the command. Ex: "ping -c 1 google.com"
        command = ['ping', param, '1', host]

        # Pinging
        return system_call(command) == 0

    ####################################
    #Fonctions concernant le dictEleves#
    #prodDictEleves                    #
    #cleanDictEleves                   #
    #refreshDictEleves                 #
    ####################################
    def prodDictEleves(self):
        """Production du dictEleves : un dict reprenant tous les élèves (obj)"""
        try:
            prog = QtWidgets.QProgressDialog()
            prog.setWindowFlags(QtCore.Qt.SplashScreen |
                                QtCore.Qt.WindowStaysOnTopHint)
            prog.setMinimumWidth(350)

            prog.setCancelButton(None)
            prog.setLabel(QtWidgets.QLabel(
                "<p>Récupération des élèves sur Smartschool</p><p>Veuillez patienter...</p>"))
            for i in range(50):
                prog.setValue(i)
                prog.show()
                # ce temps d'attente ne sert à rien mais il évite de voir la barre de progression s'afficher directement à 49%
                time.sleep(0.01)
                QtWidgets.QApplication.processEvents()

            dictEleves = {}

            webservices = "https://" + \
                self.dictConfig["urlEcole"]+"/Webservices/V3?wsdl"
            client = Client(webservices)

            # recupération de la string JSON avec toutes les infos concernant les élèves, le 1 rend la fonction récurssive
            result = client.service.getAllAccountsExtended(
                self.dictConfig['SSApiKey'], self.dictConfig['gpEleves'], 1)

            # conversion de la string json en une liste de dict
            listEleve = json.loads(result)

            # remplissage du dictEleves
            for eleve in listEleve:
                newEleve = Eleve()
                newEleve.voornaam = eleve['voornaam']
                newEleve.naam = eleve['naam']
                newEleve.internnummer = eleve['internnummer']
                newEleve.stamboeknummer = eleve['stamboeknummer']
                if self.sendComptePrincipal.isChecked() == 0:
                    newEleve.statutMsgComptePrincipal = "Non sélectionné"
                if self.sendComptesSecondaires.isChecked() == 0:
                    newEleve.statutMsgCoaccount1 = "Non sélectionné"
                    newEleve.statutMsgCoaccount2 = "Non sélectionné"
                else:
                    try:
                        newEleve.status1 = eleve['status1']
                    except KeyError:
                        newEleve.statutMsgCoaccount1 = "Pas de compte secondaire 1"
                    try:
                        newEleve.status2 = eleve['status2']
                    except KeyError:
                        newEleve.statutMsgCoaccount2 = "Pas de compte secondaire 2"

                self.dictEleves[eleve['stamboeknummer']] = newEleve
            for i in range(51):
                prog.setValue(i+50)
                QtWidgets.QApplication.processEvents()
                # ce temps d'attente ne sert à rien mais il évite de voir la barre de progression disparaitre subitement
                time.sleep(0.01)

        except Exception as e:

            # self.dialPatientez.close()
            QtWidgets.QMessageBox.warning(self, 'Erreur récupération élèves',
                                          "<p>Il n'a pas été possible de récupérer la liste des élèves sur SmartSchool. </p><p>Vérifiez la <b>configuration</b> du logiciel et refaites une tentative. </p><p>Si cette erreur persiste, contactez le développeur.</p>")
            print('Erreur : %s' % e)
            print('Message : ', traceback.format_exc())
    #

    def refreshDictEleves(self):
        """Retélécharge le dictEleves puis mets à jour la liste affichée dans le tableau"""
        self.prodDictEleves()
        self.updateFileList()
    #

    def cleanDictEleves(self):
        """Retire du dictEleves tous les élèves pour lesquels il n'y a pas ou plus de fichier à envoyer"""
        listElevesSansFichier = []
        for eleve in self.dictEleves.values():
            if eleve.filePath == "":
                listElevesSansFichier.append(eleve.stamboeknummer)
        for elem in listElevesSansFichier:
            del self.dictEleves[elem]
    #

    def changeDestinataire(self):
        """
        Modifie les valeurs de sstatutMsgComptePrincipal, statutMsgCoaccount1 et statutMsgCoaccount2 
        en fonction des valeurs des checkbox sendComptePrincipal et sendComptesSecondaires
        """
        for eleve in self.dictEleves.values():
            if self.sendComptePrincipal.isChecked():
                eleve.statutMsgComptePrincipal = "En attente"
            else:
                eleve.statutMsgComptePrincipal = "Non sélectionné"

            if self.sendComptesSecondaires.isChecked():
                eleve.statutMsgCoaccount1 = "En attente"
                eleve.statutMsgCoaccount2 = "En attente"
            else:
                eleve.statutMsgCoaccount1 = "Non sélectionné"
                eleve.statutMsgCoaccount2 = "Non sélectionné"
        self.updateFileList()

    #############################################################################
    #Fonctions concernant l'importation des fichiers vers l'interface graphique #
    #verifAvantImport                                                           #
    #importFilesDialog                                                          #
    #############################################################################
    def verifAvantImport(self):
        """Vérifie que tous est en ordre avant d'importer des fichiers"""
        self.getConfig()
        alert = "<div><p>Vous voulez importer les fichiers mais :<ul>"
        valide = 1
        if self.dictConfig["web"] == 0:
            alert += "<li>Vous ne semblez pas connecté à  <b>internet</b></li>"
            valide = 0
        if self.dictConfig["urlEcole"] == "":
            alert += "<li>Vous n'avez pas configuré l'<b>adresse SmartSchool</b> de votre école.</li>"
            valide = 0
        if self.dictConfig["SSApiKey"] == "":
            alert += "<li>Vous n'avez pas configuré la <b>clè d'accès à l'API</b>.</li>"
            valide = 0
        if self.dictConfig["gpEleves"] == 0:
            alert += "<li>Vous n'avez pas configuré le <b>nom du groupe</b> qui contient les élèves.</li>"
            valide = 0
        alert += "</ul></p></div>"

        if valide == 0:
            QtWidgets.QMessageBox.warning(
                self, 'Logiciel non configuré', alert)
        else:
            try:
                self.importFilesDialog()
            except Exception as e:
                QtWidgets.QMessageBox.warning(
                    self, 'Erreur critique', "<p>Une erreur critique s'est produite, contactez le développeur.</p>")
                print('Erreur : %s' % e)
                print('Message : ', traceback.format_exc())
    #

    def importFilesDialog(self):
        """
        Ouvre une boite de dialogue pour sélectionner les fichiers pdf
        """
        listFilesNames = QtWidgets.QFileDialog.getOpenFileNames(
            self, 'Les fichiers à envoyer', self.current_dir, filter='PDF files (*.pdf)')
        for filePath in listFilesNames[0]:
            try:

                # récupère le nom du fichier avec son extension
                fileName = os.path.basename(filePath)
                fileName = os.path.splitext(fileName)[0]
                fileName=fileName.split('_')[0]#le nom du fichier peut contenir autre chose que le matricule de l'élève mais il doit commencer par celui-ci et le reste du nom doit être séparé du matricule par le symbole "_"
                # à chaque élève du dictEleves pour lequel il y a un fichier, j'ajoute le path vers ce fichier
                self.dictEleves[fileName].filePath = filePath
            except KeyError:
                alert = "<div><p>Aucune correspondance n'a été trouvée pour le fichier avec le numéro %s</p>" % (
                    os.path.splitext(fileName)[0])
                alert += "<p>Vérifiez que ce numéro correspond à un <b>matricule élève</b>.</p>"
                alert += "<p>Vérifiez que ce numéro est convenablement encodé dans <b>Smartschool</b>.</p></div>"
                QtWidgets.QMessageBox.warning(
                    self, 'Aucune correspondance', alert)
        self.updateFileList()

    #####################################################################################
    #Fonctions concernant l'affichage de la liste des élèves à qui on envoie un message #
    #updateFileList                                                                     #
    #removeFile                                                                         #
    #####################################################################################
    def updateFileList(self):
        """
        Mets à jour le tableau présentant la liste des fichiers, le nom des élèves et le statut du message 
        """
        try:
            self.dockWidgetList.setParent(None)
        except:
            self.dockWidgetList = QtWidgets.QDockWidget(self)

        # creation du DockWidget
        self.dockWidgetList.setTitleBarWidget(QtWidgets.QLabel(
            '<p style="font-size:10pt;font-weight:bold">Fichiers</p>'))
        self.dockWidgetList.setFeatures(
            QtWidgets.QDockWidget.NoDockWidgetFeatures)
        self.dockWidgetList.setMinimumWidth(600)

        # création d'un layout vertical et d'un widget dans lequel on place le layout
        layout = QtWidgets.QVBoxLayout()
        widget = QtWidgets.QWidget()
        widget.setLayout(layout)

        # creation du tableau : paramètres
        self.tableWidget = QtWidgets.QTableWidget()
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.tableWidget.setHorizontalHeaderLabels(
            ['Nom du fichier', "Nom et prénom de l'élève", "Statut compte principal", "Statut compte sec. 1", "Statut compte sec. 2"])
        self.tableWidget.setColumnWidth(0, 120)
        self.tableWidget.setColumnWidth(1, 120)
        self.tableWidget.setColumnWidth(2, 120)
        self.tableWidget.setColumnWidth(3, 120)
        self.tableWidget.setColumnWidth(4, 120)
        # creation du tableau : contenu
        for eleve in self.dictEleves.values():
            if eleve.filePath != "":
                itemFileName = QtWidgets.QTableWidgetItem(eleve.stamboeknummer)
                nomPrenom = eleve.naam+"_"+eleve.voornaam
                itemNomPrenom = QtWidgets.QTableWidgetItem(nomPrenom)
                itemStatutMsgPrinc = QtWidgets.QTableWidgetItem(
                    eleve.statutMsgComptePrincipal)
                itemStatutMsg1 = QtWidgets.QTableWidgetItem(
                    eleve.statutMsgCoaccount1)
                itemStatutMsg2 = QtWidgets.QTableWidgetItem(
                    eleve.statutMsgCoaccount2)

                self.tableWidget.insertRow(self.tableWidget.rowCount())
                self.tableWidget.setItem(
                    self.tableWidget.rowCount()-1, 0, itemFileName)
                self.tableWidget.setItem(
                    self.tableWidget.rowCount()-1, 1, itemNomPrenom)
                self.tableWidget.setItem(
                    self.tableWidget.rowCount()-1, 2, itemStatutMsgPrinc)
                self.tableWidget.setItem(
                    self.tableWidget.rowCount()-1, 3, itemStatutMsg1)
                self.tableWidget.setItem(
                    self.tableWidget.rowCount()-1, 4, itemStatutMsg2)

        # creation du bouton remove
        self.boutton_remove = QtWidgets.QPushButton(
            QtGui.QIcon('./icons/remove.svg'), 'Retirer de la liste')
        self.boutton_remove.setMinimumHeight(30)
        self.boutton_remove.clicked.connect(self.removeFile)

        # placement des widgets
        layout.addWidget(self.tableWidget)
        layout.addWidget(self.boutton_remove)

        # placement du widget dans le dock
        self.dockWidgetList.setWidget(widget)

        # placement du dock dans la fenêtre, à droite
        self.addDockWidget(QtCore.Qt.RightDockWidgetArea, self.dockWidgetList)
    #

    def removeFile(self):
        """
        Le fichier correspondant à la ligne active est retiré de la liste des fichiers à traiter
        """
        try:
            currentRow = self.tableWidget.currentRow()
            name = self.tableWidget.item(currentRow, 0).data(0)
            self.dictEleves[name].filePath = ""
        except:
            print("Erreur dans la fonction removeFile")
            print('Erreur : %s' % e)
            print('Message : ', traceback.format_exc())

        self.updateFileList()

    ########################################
    #Fonctions concernant la configuration #
    #getConfig                             #
    #openDialConfig                        #
    #closeDialConfig                       #
    #recordValue                           #
    ########################################
    def getConfig(self):
        """Récupère les informations de configuration dans le fichier json, crée ce fichier s'il est absent, récupère d'autres paramètres de configuration"""
        try:
            confFile = open("config.json", "r")
        except FileNotFoundError:
            TempDictConfig = {}
            TempDictConfig["urlEcole"] = ""
            TempDictConfig["SSApiKey"] = ""
            TempDictConfig["interNummerExpediteur"] = ""
            TempDictConfig["gpEleves"] = ""
            myFile = open("config.json", "w")
            myFile.write(json.dumps(TempDictConfig))
            myFile.close()
            confFile = open("config.json", "r")

        confJson = json.loads(confFile.read())

        # vérification du site web de l'école
        self.dictConfig["urlEcole"] = confJson["urlEcole"]

        # vérification de l'accès au site web de SS
        # cet accès est systématiquement vérifé à chaque fois qu'on fait appel à checkConfig
        hostname = "https://"+self.dictConfig["urlEcole"]
        response = os.system("ping -c 1 " + hostname)
        self.dictConfig["web"] = str(response)  # 0 down 1 up

        # vérification de la clé d'accès à l'APi_SS
        self.dictConfig["SSApiKey"] = confJson["SSApiKey"]

        # vérification du internummer
        self.dictConfig["interNummerExpediteur"] = confJson["interNummerExpediteur"]

        # vérification du gpEleves
        self.dictConfig["gpEleves"] = confJson["gpEleves"]

        # vérification du dictEleves, production si nécessaire
        # le dict élèves est long à produire et il n'est pas nécessaire de le produire à chaque appel à checkConfig
        # il sera produit une fois lorsque le programme commence et il ne sera produit après que si il est vide
        if self.dictEleves == {}:  # si le dict est vide, on le produit
            self.prodDictEleves()
    #

    def openDialConfig(self):
        """
        Ouvre une boite de dialogue qui permet de configurer le logiciel ou de modifier les valeurs de configuration
        """
        try:  # on commence par récupérer les données de configuration sachant qu'il est possible que cela ne fonctionne pas
            self.getConfig()
        except:
            pass

        # ouvre la boite de dialogue
        self.dialConfig = QtWidgets.QDialog(parent=gui)
        mon_layout = QtWidgets.QVBoxLayout()  # création d'un layout vertical

        mon_label = QtWidgets.QLabel("<p>Url du site SS de l'école</p>")
        mon_layout.addWidget(mon_label)

        self.saisieUrl = QtWidgets.QLineEdit()
        if self.dictConfig["urlEcole"] != "":
            self.saisieUrl.setText(self.dictConfig["urlEcole"])
        else:
            self.saisieUrl.setPlaceholderText("Ex. : mon_école.smartschool.be")
        mon_layout.addWidget(self.saisieUrl)

        mon_label = QtWidgets.QLabel("<p>Clé d'accès à l'API</p>")
        mon_layout.addWidget(mon_label)

        self.saisieApiKey = QtWidgets.QLineEdit()
        if self.dictConfig["SSApiKey"] != "":
            self.saisieApiKey.setText(self.dictConfig["SSApiKey"])
        else:
            self.saisieApiKey.setPlaceholderText(
                "Le code d'accèsà l'API de SmartSchool")
        mon_layout.addWidget(self.saisieApiKey)

        mon_label = QtWidgets.QLabel(
            "<p>Numéro interne SS de l'expéditeur</p>")
        mon_layout.addWidget(mon_label)

        self.saisieInternummer = QtWidgets.QLineEdit()
        if self.dictConfig["interNummerExpediteur"] != "":
            self.saisieInternummer.setText(
                self.dictConfig["interNummerExpediteur"])
        else:
            self.saisieInternummer.setPlaceholderText(
                "Numéro interne SmartSchool de l'expéditeur")
        mon_layout.addWidget(self.saisieInternummer)

        mon_label = QtWidgets.QLabel(
            "<p>Groupe des élèves dans SmartSchool</p>")
        mon_layout.addWidget(mon_label)

        self.saisieGpEleves = QtWidgets.QLineEdit()
        if self.dictConfig["gpEleves"] != "":
            self.saisieGpEleves.setText(self.dictConfig["gpEleves"])
        else:
            self.saisieGpEleves.setPlaceholderText("Ex. Elèves")
        mon_layout.addWidget(self.saisieGpEleves)

        mon_bouton = QtWidgets.QPushButton('Enregistrer les valeurs')
        mon_bouton.clicked.connect(self.recordValue)
        mon_layout.addWidget(mon_bouton)

        mon_bouton = QtWidgets.QPushButton('Fermer sans enregistrer')
        mon_bouton.clicked.connect(self.closeDialConfig)
        mon_layout.addWidget(mon_bouton)

        # ajout du layout dans la fenetre
        self.dialConfig.setLayout(mon_layout)
        self.dialConfig.exec_()  # affichage de la boite
    #

    def closeDialConfig(self):
        self.dialConfig.close()
    #

    def recordValue(self):
        """
        Enregistre les valeurs de configuration dans un fichier JSON
        """
        TempDictConfig = {}
        TempDictConfig["urlEcole"] = str(self.saisieUrl.text())
        TempDictConfig["SSApiKey"] = str(self.saisieApiKey.text())
        TempDictConfig["interNummerExpediteur"] = str(
            self.saisieInternummer.text())
        TempDictConfig["gpEleves"] = str(self.saisieGpEleves.text())
        myFile = open("config.json", "w")
        myFile.write(json.dumps(TempDictConfig))
        myFile.close()
        self.getConfig()  # apres avoir enregistré les valeurs de config dansle fichier JSON, on génère le dictConfig
        self.dialConfig.close()

    ############################################
    #Fonctions concernant l'envoi des messages #
    #verifAvantEnvoi                           #
    #sendMsg                                   #
    #sendAllFiles                              #
    ############################################
    def verifAvantEnvoi(self):
        """Vérifie que tout est en ordre avant d'envoyer les fichiers"""
        self.getConfig()  # je refais un getConfig ici car après un permier envoie le dictEleves est vidé et il faut le recharger depuis SS. La fonction getConfig recharge le dictEleves si il est vide

        noFile = 1
        for eleve in self.dictEleves.values():  # on vérifie s'il y a au moins un fichier à envoyer
            if eleve.filePath != "":
                noFile = 0

        alert = "<div><p>Vous voulez envoyer les messages mais :<ul>"
        valide = 1
        if noFile:
            alert += "<li>Vous n'avez importé <b>aucun fichier</b> à envoyer</li>"
            valide = 0
        if self.dictConfig["web"] == 0:
            alert += "<li>Vous ne semblez pas connecté à  <b>internet</b></li>"
            valide = 0
        if self.dictConfig["urlEcole"] == "":
            alert += "<li>Vous n'avez pas configuré l'<b>adresse SmartSchool</b> de votre école.</li>"
            valide = 0
        if self.dictConfig["SSApiKey"] == "":
            alert += "<li>Vous n'avez pas configuré la<b>clè d'accès à l'API</b>.</li>"
            valide = 0
        if self.dictConfig["interNummerExpediteur"] == "":
            alert += "<li>Vous n'avez pas configuré le <b>numéro interne</b> de l'expéditeur.</li>"
            valide = 0
        if self.dictConfig["gpEleves"] == 0:
            alert += "<li>Vous n'avez pas configuré le <b>nom du groupe</b> qui contient les élèves.</li>"
            valide = 0
        if self.sujet.text() == "":  # or self.message.toHtml()=="":
            valide = 0
            alert += "<li>Vous avez oublié le <b>sujet</b></li>"
        if self.message.toPlainText() == "":
            valide = 0
            alert += "<li>Vous avez oublié le <b>message</b></li>"
        if self.sendComptePrincipal.isChecked() == 0 & self.sendComptesSecondaires.isChecked() == 0:
            valide = 0
            alert += "<li>Vous n'avez sélectionné <b>aucun compte</b> pour l'envoi</li>"

        alert += "</ul></p></div>"

        if valide == 0:
            QtWidgets.QMessageBox.warning(self, 'Envoi impossible', alert)
        else:
            try:
                self.sendAllFiles()
            except Exception as e:
                QtWidgets.QMessageBox.warning(
                    self, 'Erreur critique', "<p>Une erreur critique s'est produite, contactez le développeur.</p>")
                print('Erreur : %s' % e)
                print('Message : ', traceback.format_exc())
    #

    def sendMsg(self, eleve):
        """envoie le fichier fileName(str) à l'élève eleve(obj) """
        webservices = "https://" + \
            self.dictConfig["urlEcole"]+"/Webservices/V3?wsdl"
        client = Client(webservices)

        attachmentName = eleve.naam+"_"+eleve.voornaam+".pdf"
        with open(eleve.filePath, "rb") as myFile:
            encodedFile = base64.b64encode(myFile.read())
        encodedFile = encodedFile.decode()  # on recupere la partie string du byte
        attachment = [{"filename": attachmentName, "filedata": encodedFile}]

        jsonAttachment = json.dumps(attachment)

        fromaddr = self.dictConfig["interNummerExpediteur"]
        toaddr = eleve.internnummer
        subject = self.sujet.text()

        body = self.message.toHtml()
        if self.sendComptePrincipal.isChecked():
            result0 = client.service.sendMsg(
                self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 0, 0)
            if result0 == 0:
                eleve.statutMsgComptePrincipal = "Envoyé"
            else:
                eleve.statutMsgComptePrincipal = "Échec"
        else:
            result0 = 0  # s'il n'y a pas de message à envoyer au compte principal, il n'y a donc pas d'erreur liée à cet envoi

        if self.sendComptesSecondaires.isChecked():
            if eleve.status1 == "actief":
                result1 = client.service.sendMsg(
                    self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 1, 0)
                if result1 == 0:
                    eleve.statutMsgCoaccount1 = "Envoyé"
                else:
                    eleve.statutMsgCoaccount1 = "Échec"
            else:
                result1 = 0  # s'il n'y a pas de compte secondaire 1, il n'y a pas d'erreur liée à l'envoi sur ce compte
            if eleve.status2 == "actief":
                result2 = client.service.sendMsg(
                    self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 2, 0)
                if result2 == 0:
                    eleve.statutMsgCoaccount2 = "Envoyé"
                else:
                    eleve.statutMsgCoaccount2 = "Échec"
            else:
                result2 = 0  # s'il n'y a pas de compte secondaire 2, il n'y a pas d'erreur liée à l'envoi sur ce compte
        else:
            result1 = 0  # s'il n'y a pas de message à envoyer aux comptes secondaires, il n'y a donc pas d'erreur liée à ces envois
            result2 = 0  # idem

        # si aucun problème n'a été rencontré dans l'envoi des messages
        if result0 == 0 & result1 == 0 & result2 == 0:
            # get the path to the file in the current directory
            # on récupère le path du fichier
            src = os.path.dirname(eleve.filePath)
            # rename the original file
            # on produit le nouveau nom du fichier
            newName = os.path.join(src, attachmentName)
            os.rename(eleve.filePath, newName)  # on renomme le fichier
            eleve.filePath = ""  # on supprime le fichier à envoyer de l'objet Eleve

        self.updateFileList()
        QtWidgets.QApplication.processEvents()
    #

    def sendAllFiles(self):
        """Envoi tous les fichiers importés"""
        valide = 1
        alert = "<div><p>Vous voulez envoyer les messages mais :<ul>"
        if self.sujet.text() == "":  # or self.message.toHtml()=="":
            valide = 0
            alert += "<li>Vous avez oublié le <b>sujet</b></li>"
        if self.message.toPlainText() == "":
            valide = 0
            alert += "<li>Vous avez oublié le <b>message</b></li>"
        alert += "</ul></p></div>"

        if valide == 0:
            QtWidgets.QMessageBox.warning(
                self, 'Pas de sujet - pas de message', alert)
        else:
            self.cleanDictEleves()
            for eleve in self.dictEleves.values():
                self.sendMsg(eleve)
        self.cleanDictEleves()
        if self.dictEleves == {}:
            message = "<div><p>Tous les documents ont été envoyés avec succès.</p><p>Si vous importez de nouveaux fichiers la liste des élèves sera à nouveau importée.</p></div>"
            QtWidgets.QMessageBox.information(self, 'Envoi terminé', message)

    #########################################################
    #Fonctions concernant directement l'interface graphique #
    #about                                                  #
    #aide                                                   #
    #appExit                                                #
    #########################################################
    def about(self):
        """
        Affiche une boite d'information sur ce logiciel : à quoi sert-il, qui l'a écris, avec quels langages
        """
        message = '<div><p><b>SSMsg</b> est un outil '
        message += "permettant d'envoyer des fichiers par lots. L'envoi se fait  sur la plateforme SmartSchool via un message accompagné d'une pièce jointe.</p>"
        message += "<p>Ce programme renomme également les fichiers en leur donnant le nom et le prénom de l'élève concerné par le fichier.</p><p> <b>Auteur : </b>J.N. Gautier</p>"
        message += "<p> <b>Language : </b>Python %s</p>  <p> <b>Interface : </b>Qt %s</p> <p> Pour signaler un bug ou proposer une amélioration, <br/>" % (
            platform.python_version(), QtCore.QT_VERSION_STR)
        message += 'vous pouvez me contacter sur Github.</p></div>'
        QtWidgets.QMessageBox.about(self, "A propos de SSMsg", message)
    #

    def aide(self):
        """
        Affiche une boite d'information sur ce qu'il faut faire en cas de problème
        """
        message = ''
        message += '<div><p>Le profil du service Web utilisé pour cette application doit être configuré avec au minimum les méthodes autorisées :'
        message += "<ul><li>getAllAccountsExtended</li><li> sendMsg</li></ul></p>"
        message += "<p>Pour plus d'informations  sur les profils de service web et les méthodes autorisées, veuillez vous référer à la <b>documentation de SmartSchool</b>.</p>"
        message += "<p>Le nom des fichiers à envoyer doit correspondre au matricule de l'élève, encodé dans le champ <b>stamboeknummer</b> de SmartSchool.</p>"
        message += "<div><p>Pour obtenir de l'aide conçernant l'emploi de ce logiciel "
        message += "vous pouvez me contacter sur Github.</p></div>"
        QtWidgets.QMessageBox.information(self, 'Aide', message)
    #

    def appExit(self):
        """
        Ferme le logiciel
        """
        print("Au revoir!")
        app.quit()


class Eleve():
    def __init__(self):
        self.voornaam = ""
        self.naam = ""
        self.internnummer = ""
        self.stamboeknummer = ""
        self.filePath = ""
        self.status1 = ""
        self.status2 = ""
        self.statutMsgComptePrincipal = "En attente"
        self.statutMsgCoaccount1 = "En attente"
        self.statutMsgCoaccount2 = "En attente"


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    gui = Gui()
    gui.getConfig()
    gui.show()
    app.exec_()
