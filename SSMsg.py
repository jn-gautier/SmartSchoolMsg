#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import base64
import json
import traceback
import platform
import csv

from zeep import Client  # doit être installé en plus de python à l'aide de pip

from platform import system as system_name  # Returns the system/OS name

from PyQt5 import QtGui
from PyQt5 import QtCore
from PyQt5 import QtWidgets


class Gui(QtWidgets.QMainWindow):

    def __init__(self):
        super(Gui, self).__init__()
        self.dictConfig = {}
        self.dictUtilisateurs = {}
        self.dictError = {}
        self.current_dir = os.path.expanduser("~")
        self.premierAffichage=True

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

        # refreshDictAction : recharge le dictUtilisateurs depuis SS
        refreshAction = QtWidgets.QAction(QtGui.QIcon(
            './icons/refresh.svg'), "Recharger liste utilisateurs", self)
        refreshAction.setStatusTip(
            "Recharger le dictionnaire depuis SmartSchool")
        refreshAction.triggered.connect(self.refreshDictUtilisateurs)

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
        dockWidgetMessage.setMaximumWidth(400)

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
        ##################
        #le Dock Renommer#
        ##################
        # création du Dock
        dockWidgetRename = QtWidgets.QDockWidget(self)
        dockWidgetRename.setTitleBarWidget(QtWidgets.QLabel(
            '<p style="font-size:11pt;font-weight:bold">Renommer</p>'))
        dockWidgetRename.setFeatures(
            QtWidgets.QDockWidget.NoDockWidgetFeatures)

        # création d'un layout vertical et d'un widget dans lequel on place le layout
        layout = QtWidgets.QVBoxLayout()
        widget = QtWidgets.QWidget()
        widget.setLayout(layout)

        # création des widget du message
        self.Renommer = QtWidgets.QCheckBox("Renommer")
        self.Renommer.setToolTip(
            "<p>Renomme les fichiers lorsqu'ils sont envoyés</p>")
        self.Renommer.setCheckState(2)
        
        self.renamePattern = QtWidgets.QLineEdit()
        self.renamePattern.setText("%nom ; %prenom")
        
        # placement des widget de renommage dans le layout
        layout.addWidget(self.Renommer)
        layout.addWidget(QtWidgets.QLabel("<p><b>Motif</b></p>"))
        layout.addWidget(self.renamePattern)

        # placement du widget dans le dock
        dockWidgetRename.setWidget(widget)

        # placement du dock dans la fenêtre, à gauche
        self.addDockWidget(QtCore.Qt.LeftDockWidgetArea, dockWidgetRename)

    def SSApiClient(self):
        """Crée un client pour l'APi_SS et renvoie celui-ci"""
        webservices = "https://" + self.dictConfig["urlEcole"] + "/Webservices/V3?wsdl"
        try:
            client = Client(webservices)
        except Exception as e:
            print('Erreur : %s' % e)
        return client
    
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
    
    def getErrorCodeMessage(self):
        """Création du dictError, un dict qui a comme clé les codes d'erreur et comme valeur la signification de ces codes"""
        
        try:
            self.client
        except:
            self.client=self.SSApiClient()
        
        #recupération des codes d'erreurs et production du dict pour interpréter ces codes
        result =self.client.service.returnJsonErrorCodes()
        self.dictError=json.loads(result)
       

    ####################################
    #Fonctions concernant le dictUtilisateurs#
    #prodDictUtilisateurs                    #
    #cleanDictUtilisateurs                   #
    #refreshDictUtilisateurs                 #
    ####################################
    def getListUtilisateurs(self):
        """
        Récupère les liste de tous les utilisateurs de la plateforme avec toutes les nformations sur ces utilisateurs, renvoie une liste dont chaque élément est un dict
        """
        try:
            self.client
        except:
            self.client=self.SSApiClient()
        
            
        # recupération de la string JSON avec toutes les infos concernant les utilisateurs, le 1 rend la fonction récurssive,la string vide comme deuxième argument a comme conséquence qu'on télécharge touts les utilisateurs de la plateforme et non un sous groupe comme "prof" ou "élèves"
        result = self.client.service.getAllAccountsExtended(self.dictConfig['SSApiKey'], '', 1)
        
        # conversion de la string json en une liste de dict
        try:
            listUtilisateurs = json.loads(result)
        except TypeError:
            if isinstance(result, int):
                print("Le fichier reçu n'est pas un fichier JSON.")
                print("Vous avez reçu un message d'erreur en provenance de l'API-SS.")
                print("Message d'erreur reçu :", self.dictError[str(result)])
                msgError="<p>Il n'a pas été possible de récupérer la liste des utilisateurs sur SmartSchool. </p><p>Vérifiez la <b>configuration</b> du logiciel et refaites une tentative. </p><p>Si cette erreur persiste, contactez le développeur.</p><p>Vous avez reçu un message d'erreur en provenance de l'API-SS.</p><p>Message d'erreur reçu :  %s </p>" %self.dictError[str(result)]
                QtWidgets.QMessageBox.warning(self, 'Erreur récupération utilisateurs',msgError)
        #print(listUtilisateurs)
        return listUtilisateurs
        
    def prodDictUtilisateurs(self):
        """Production du dictUtilisateurs : un dict reprenant tous les utilisateurs (obj)"""
        try:
            prog = QtWidgets.QProgressDialog()
            prog.setWindowFlags(QtCore.Qt.SplashScreen |
                                QtCore.Qt.WindowStaysOnTopHint)
            prog.setMinimumWidth(350)

            prog.setCancelButton(None)
            prog.setLabel(QtWidgets.QLabel(
                "<p>Récupération des utilisateurs sur Smartschool</p><p>Veuillez patienter...</p>"))
            for i in range(50):
                prog.setValue(i)
                prog.show()
                # ce temps d'attente ne sert à rien mais il évite de voir la barre de progression s'afficher directement à 49%
                time.sleep(0.01)
                QtWidgets.QApplication.processEvents()

            dictUtilisateurs = {}

            
            self.listUtilisateurs=self.getListUtilisateurs()
            
            # remplissage du dictUtilisateurs
            for utilisateur in self.listUtilisateurs:
                newUtilisateur = Utilisateur()
                newUtilisateur.voornaam = utilisateur['voornaam']
                newUtilisateur.naam = utilisateur['naam']
                newUtilisateur.internnummer = utilisateur['internnummer']
                newUtilisateur.stamboeknummer = utilisateur['stamboeknummer']
                try: #on tente de trouver le champ d'identification des utilisateurs donné par l'utilisateur du logiciel. Si ce champ n'existe pas ou n'a pas été défini on mets le stamboeknummer. l'existence de ce champ sera testée lorsqu'on modifie la configuration
                    newUtilisateur.champIdentSS = utilisateur[self.dictConfig['champIdentSS']]
                except KeyError:
                    newUtilisateur.champIdentSS = utilisateur['stamboeknummer']
                if self.sendComptePrincipal.isChecked() == 0:
                    newUtilisateur.statutMsgComptePrincipal = "Non sélectionné"
                if self.sendComptesSecondaires.isChecked() == 0:
                    newUtilisateur.statutMsgCoaccount1 = "Non sélectionné"
                    newUtilisateur.statutMsgCoaccount2 = "Non sélectionné"
                else:
                    try:
                        newUtilisateur.status1 = utilisateur['status1']
                    except KeyError:
                        newUtilisateur.statutMsgCoaccount1 = "Pas de compte secondaire 1"
                    try:
                        newUtilisateur.status2 = utilisateur['status2']
                    except KeyError:
                        newUtilisateur.statutMsgCoaccount2 = "Pas de compte secondaire 2"
                
                if self.doesChampIdentSSExist():
                    champIdentSS=self.dictConfig['champIdentSS']
                    self.dictUtilisateurs[utilisateur[champIdentSS]] = newUtilisateur
                else:
                    self.dictUtilisateurs[utilisateur['stamboeknummer']] = newUtilisateur
                
            for i in range(51):
                prog.setValue(i+50)
                QtWidgets.QApplication.processEvents()
                # ce temps d'attente ne sert à rien mais il évite de voir la barre de progression disparaitre subitement
                time.sleep(0.01)
        
        
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, 'Erreur création du dictUtilisateurs',
                                          "<p>Il n'a pas été possible de créer le dictionnaire des utilisateurs. </p><p>Vérifiez la <b>configuration</b> du logiciel et refaites une tentative. </p><p>Si cette erreur persiste, contactez le développeur.</p>")
            print('Erreur : %s' % e)
            print('Message : ', traceback.format_exc())
    #
    def refreshDictUtilisateurs(self):
        """Retélécharge le dictUtilisateurs puis mets à jour la liste affichée dans le tableau"""
        self.prodDictUtilisateurs()
        self.updateFileList()
    #
    def cleanDictUtilisateurs(self):
        """Retire du dictUtilisateurs tous les utilisateurs pour lesquels il n'y a pas ou plus de fichier à envoyer"""
        listUtilisateursSansFichier = []
        for utilisateur in self.dictUtilisateurs.values():
            if utilisateur.filePath == "":
                listUtilisateursSansFichier.append(utilisateur.champIdentSS)
        for elem in listUtilisateursSansFichier:
            del self.dictUtilisateurs[elem]
    #
    def changeDestinataire(self):
        """
        Modifie les valeurs de statutMsgComptePrincipal, statutMsgCoaccount1 et statutMsgCoaccount2 
        en fonction des valeurs des checkbox sendComptePrincipal et sendComptesSecondaires
        """
        for utilisateur in self.dictUtilisateurs.values():
            if self.sendComptePrincipal.isChecked():
                utilisateur.statutMsgComptePrincipal = "En attente"
            else:
                utilisateur.statutMsgComptePrincipal = "Non sélectionné"

            if self.sendComptesSecondaires.isChecked():
                utilisateur.statutMsgCoaccount1 = "En attente"
                utilisateur.statutMsgCoaccount2 = "En attente"
            else:
                utilisateur.statutMsgCoaccount1 = "Non sélectionné"
                utilisateur.statutMsgCoaccount2 = "Non sélectionné"
        self.updateFileList()

    #############################################################################
    #Fonctions concernant l'importation des fichiers vers l'interface graphique #
    #verifAvantImport                                                           #
    #importFilesDialog                                                          #
    #############################################################################
    def verifAvantImport(self):
        """Vérifie que tous est en ordre avant d'importer des fichiers"""
        valid , message = self.isConfigValid()
        alert = "<div><p>Vous voulez importer les fichiers mais :<ul>"
        alert += message
        alert += "</ul></p></div>"

        if not valid :
            QtWidgets.QMessageBox.warning(self, 'Logiciel non configuré', alert)
        else:
            #production du dictConfig
            self.getConfig()
            #production du dictError
            self.getErrorCodeMessage()
            # vérification du dictUtilisateurs, production si nécessaire
            if self.dictUtilisateurs == {}:  # si le dict est vide, on le produit
                self.prodDictUtilisateurs()
            try:
                self.importFilesDialog()
            except Exception as e:
                QtWidgets.QMessageBox.warning(
                    self, 'Erreur critique', "<p>Une erreur critique s'est produite lors de l'importation des documents, contactez le développeur.</p>")
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
                fileName = os.path.splitext(fileName)[0]#on récupère le nom sans l'extension
                fileName=fileName.split('_')[0]#le nom du fichier peut contenir autre chose que le matricule de l'utilisateur mais il doit commencer par celui-ci et le reste du nom doit être séparé du matricule par le symbole "_"
                # à chaque utilisateur du dictUtilisateurs pour lequel il y a un fichier, j'ajoute le path vers ce fichier
                self.dictUtilisateurs[fileName].filePath = filePath
            except KeyError:
                alert = "<div><p>Aucune correspondance n'a été trouvée pour le fichier avec le numéro %s</p>" % (
                    os.path.splitext(fileName)[0])
                alert += "<p>Vérifiez que ce numéro correspond à la <b>valeur du champ de correspondance</b> pour cet utilisateur</b>.</p>"
                alert += "<p>Vérifiez que ce numéro est convenablement encodé dans <b>Smartschool</b>.</p></div>"
                QtWidgets.QMessageBox.warning(
                    self, 'Aucune correspondance', alert)
        self.updateFileList()

    #####################################################################################
    #Fonctions concernant l'affichage de la liste des utilisateurs à qui on envoie un message #
    #updateFileList                                                                     #
    #removeFile                                                                         #
    #####################################################################################
    def updateFileList(self):
        """
        Mets à jour le tableau présentant la liste des fichiers, le nom des utilisateurs et le statut du message 
        """
        try:
            self.dockWidgetList.setParent(None)
        except:
            self.dockWidgetList = QtWidgets.QDockWidget(self)
        if self.premierAffichage:#la fenêtre s'affcihe en plein écran lors du premier affichage, après si l'utilisateur à décidé de la réduire on la laisse réduite
            self.showMaximized()
            self.premierAffichage=False
        # creation du DockWidget
        self.dockWidgetList.setTitleBarWidget(QtWidgets.QLabel(
            '<p style="font-size:10pt;font-weight:bold">Fichiers</p>'))
        self.dockWidgetList.setFeatures(
            QtWidgets.QDockWidget.NoDockWidgetFeatures)
        self.dockWidgetList.setMinimumWidth(800)
        
        # création d'un layout vertical et d'un widget dans lequel on place le layout
        layout = QtWidgets.QVBoxLayout()
        widget = QtWidgets.QWidget()
        widget.setLayout(layout)

        # creation du tableau : paramètres
        self.tableWidget = QtWidgets.QTableWidget()
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setEditTriggers(QtWidgets.QTableWidget.NoEditTriggers)
        self.tableWidget.setHorizontalHeaderLabels(
            ['Nom du fichier', "Nom et prénom de l'utilisateur", "Statut compte principal", "Statut compte sec. 1", "Statut compte sec. 2"])
        
        self.tableWidget.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(1, QtWidgets.QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(2, QtWidgets.QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(3, QtWidgets.QHeaderView.Stretch)
        self.tableWidget.horizontalHeader().setSectionResizeMode(4, QtWidgets.QHeaderView.Stretch)
        # creation du tableau : contenu
        for utilisateur in self.dictUtilisateurs.values():
            if utilisateur.filePath != "":
                itemFileName = QtWidgets.QTableWidgetItem(os.path.basename(utilisateur.filePath))
                nomPrenom = utilisateur.naam+"_"+utilisateur.voornaam
                itemNomPrenom = QtWidgets.QTableWidgetItem(nomPrenom)
                itemStatutMsgPrinc = QtWidgets.QTableWidgetItem(
                    utilisateur.statutMsgComptePrincipal)
                itemStatutMsg1 = QtWidgets.QTableWidgetItem(
                    utilisateur.statutMsgCoaccount1)
                itemStatutMsg2 = QtWidgets.QTableWidgetItem(
                    utilisateur.statutMsgCoaccount2)

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
        

    def removeFile(self):
        """
        Le fichier correspondant à la ligne active est retiré de la liste des fichiers à traiter
        """
        try:
            currentRow = self.tableWidget.currentRow()
            name = self.tableWidget.item(currentRow, 0).data(0)
            for utilisateur in self.dictUtilisateurs.values():
                if os.path.basename(utilisateur.filePath)==name:
                    utilisateur.filePath = ""
        except AttributeError : #s'il n'y a pas de ligne dans le tableau
            pass
        except Exception as e:
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
        """Récupère les informations de configuration dans le fichier json, crée ce fichier s'il est absent, donne une valeur vide aux paramètres qui ne seraient pas présents"""
        try:
            confFile = open("config.json", "r")
        except FileNotFoundError:
            TempDictConfig = {}
            TempDictConfig["urlEcole"] = ""
            TempDictConfig["SSApiKey"] = ""
            TempDictConfig["interNummerExpediteur"] = ""
            TempDictConfig["champIdentSS"] = ""
            myFile = open("config.json", "w")
            myFile.write(json.dumps(TempDictConfig))
            myFile.close()
            confFile = open("config.json", "r")

        confJson = json.loads(confFile.read())

        # vérification du site web de l'école
        try:
            self.dictConfig["urlEcole"] = confJson["urlEcole"]
        except KeyError:
            self.dictConfig["urlEcole"] = ""

        # vérification de l'accès au site web de SS
        # cet accès est systématiquement vérifé à chaque fois qu'on fait appel à checkConfig
        #hostname = "https://"+self.dictConfig["urlEcole"]
        #response = os.system("ping -c 1 " + hostname)
        #self.dictConfig["web"] = str(response)  # 0 down 1 up

        # vérification de la clé d'accès à l'APi_SS
        try:
            self.dictConfig["SSApiKey"] = confJson["SSApiKey"]
        except KeyError:
            self.dictConfig["SSApiKey"] = ""

        # vérification du internummer
        try:
            self.dictConfig["interNummerExpediteur"] = confJson["interNummerExpediteur"]
        except KeyError:
            self.dictConfig["interNummerExpediteur"] = ""
         
        # vérification du champIdentSS
        try:
            self.dictConfig["champIdentSS"] = confJson["champIdentSS"]
        except KeyError:
            self.dictConfig["champIdentSS"] = ""
        
        
    #
    def isConfigValid(self):
        """
        Vérifie si les différents paramètres possèdent une valeur, pas si cette valeur est la bonne.
        En fonction des paramètres qui ne sont pas configurés, un message d'erreur est généré.
        La fonction renvoie True si les paramètres sont configurés, False sinon. 
        Elle renvoie aussi le message d'erreur qui n'est jamais une string vide car il y a au minimum le début du message.
        """
        configValide=True
        #alert = "<div><p>Votre configuration n'est pas valide :<ul>"
        alert=''
        
        valid , message = self.isUrlEcoleValid()
        alert+=message
        if valid==False :
            configValide=False
            
        valid , message = self.isSSApiKeyValid()
        alert+=message
        if valid==False :
            configValide=False
            
        valid , message = self.isInterNummerExpediteurValid()
        alert+=message
        if valid==False :
            configValide=False
        
        valid , message = self.doesChampIdentSSExist()
        alert+=message
        if valid==False :
            configValide=False
        else: #on ne teste l'unicité du champIdentSS que s'il existe
            valid , message = self.isChampIdentSSUnique()
            alert+=message
            if valid==False :
                configValide=False
        return configValide, alert

    def isUrlEcoleValid(self):
        """
        Renvoie True si UrlEcole possède une valeur, sinon renvoie False
        Renvoie aussi un message d'erreur qui est vide lorsque la config est valide.
        """
        valid=True
        alert=''
        if self.dictConfig["urlEcole"] == "" :
            valid=False
            alert = "<li>Vous n'avez pas configuré l'<b>adresse SmartSchool</b> de votre école.</li>"
        return valid, alert
    
    def isSSApiKeyValid(self):
        """
        Renvoie True si SSApiKey possède une valeur, sinon renvoie False.
        Renvoie aussi un message d'erreur qui est vide lorsque la config est valide.
        """
        valid=True
        alert=''
        if self.dictConfig["SSApiKey"] == "" :
            valid=False
            alert = "<li>Vous n'avez pas configuré la <b>clè d'accès à l'API</b>.</li>"
        return  valid , alert

    def isInterNummerExpediteurValid(self):
        """
        Renvoie True si interNummerExpediteur possède une valeur, sinon renvoie False.
        Renvoie aussi un message d'erreur qui est vide lorsque la config est valide.
        """
        valid=True
        alert=''
        if self.dictConfig["interNummerExpediteur"] == "" :
            valid=False
            alert = "<li>Vous n'avez pas configuré le <b>numéro interne</b> de l'expéditeur.</li>"
        return  valid , alert
    
    def isChampIdentSSUnique(self):
        """
        Renvoie False s'il existe au moins deux utilisateurs qui possèdent la même valeur pour le champ d'identification, sinon renvoie True.
        Renvoie aussi un message d'erreur qui est vide lorsque le champ est unique.
        """
        listValeurChampIdentSS=[]
        listeValeursNonUniques=[]
        champIdentSS=self.dictConfig['champIdentSS']
        for utilisateur in self.listUtilisateurs:
            if utilisateur[champIdentSS] in listValeurChampIdentSS: #on forme une liste avec les valeurs présentes plus d'une fois
                listeValeursNonUniques.append(utilisateur[champIdentSS])
            if (utilisateur[champIdentSS] is not None) & (utilisateur[champIdentSS] != 'NULL'): #les utilisateurs avec une valeur NULL ou sans valeur sont ignorés
                listValeurChampIdentSS.append (utilisateur[champIdentSS])
        listeValeursNonUniques=set(listeValeursNonUniques)
        if len(listValeurChampIdentSS)==len(set(listValeurChampIdentSS)): #le champ est unique s'il n'y a pas deux utilisateurs ayant la même valeur pour ce champ
            valid=True
            alert=''
        else:
            valid=False
            alert = "<li>Le <b>champ de SmartSchool utilisé pour identifier les utilisateurs </b> n'est pas unique : il possède la même valeur pour des utilisateurs différents."
            alert += "<br>La ou les valeurs présentes plus d'une fois sont : <ul>"
            for valeur in listeValeursNonUniques:
                alert += "<li> %s</li>" %valeur
            alert += "</ul></li>"
        return valid , alert
    
    def doesChampIdentSSExist(self):
        valid=True
        alert=''
        champIdentSS=self.dictConfig['champIdentSS']
        for utilisateur in self.listUtilisateurs:
            try:
                utilisateur[champIdentSS]
            except:
                valid=False
        if valid==False:
            alert="<li>Le <b>champ de SmartSchool utilisé pour identifier les utilisateurs </b> n'existe pas dans le profil d'au moins un utilisateur sur la plateforme SmartSchool de votre établissement.</li>"
        return valid , alert
        
    
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
        
        #Url du site de l'école : label
        mon_label = QtWidgets.QLabel("<p>Url du site SS de l'école</p>")
        mon_layout.addWidget(mon_label)
        
        #Url du site de l'école : boite de saisie
        self.saisieUrl = QtWidgets.QLineEdit()
        if self.dictConfig["urlEcole"] != "":
            self.saisieUrl.setText(self.dictConfig["urlEcole"])
        else:
            self.saisieUrl.setPlaceholderText("Ex. : mon_école.smartschool.be")
        mon_layout.addWidget(self.saisieUrl)
        
        #Clé d'accès à l'API : label
        mon_label = QtWidgets.QLabel("<p>Clé d'accès à l'API</p>")
        mon_layout.addWidget(mon_label)
        
        #Clé d'accès à l'API : boite de saisie
        self.saisieApiKey = QtWidgets.QLineEdit()
        if self.dictConfig["SSApiKey"] != "":
            self.saisieApiKey.setText(self.dictConfig["SSApiKey"])
        else:
            self.saisieApiKey.setPlaceholderText(
                "Le code d'accès à l'API de SmartSchool")
        mon_layout.addWidget(self.saisieApiKey)
        
        #Numéro interne de l'expéditeur : label
        mon_label = QtWidgets.QLabel(
            "<p>Numéro interne SS de l'expéditeur</p>")
        mon_layout.addWidget(mon_label)
        
        #Numéro interne de l'expéditeur : boite de saisie
        self.saisieInternummer = QtWidgets.QLineEdit()
        if self.dictConfig["interNummerExpediteur"] != "":
            self.saisieInternummer.setText(
                self.dictConfig["interNummerExpediteur"])
        else:
            self.saisieInternummer.setPlaceholderText(
                "Numéro interne SmartSchool de l'expéditeur")
        mon_layout.addWidget(self.saisieInternummer)
        
        
        #Champ d'identification dans SS : label
        mon_label = QtWidgets.QLabel(
            "<p>Champ de SmartSchool utilisé pour identifier les utilisateurs</p>")
        mon_layout.addWidget(mon_label)
        
        #Champ d'identification dans SS : boite de saisie
        self.saisieChampCorrespondanceSS = QtWidgets.QLineEdit()
        if self.dictConfig["champIdentSS"] != "":
            self.saisieChampCorrespondanceSS.setText(self.dictConfig["champIdentSS"])
        else:
            self.saisieChampCorrespondanceSS.setPlaceholderText("Normalement : stamboeknummer ou internnummer")
        myTip="<p>Le nom des fichiers à envoyer doit contenir une valeur permettant d'identifier son destinataire.</p>"
        myTip+="<p>Il s'agit normalement de la valeur de son <b>stamboeknummer</b> ou de son <b>internnummer</b> mais un autre champ de correspondance pourrait être utilisé.</p>"
        myTip+="<p>Ce champ doit exister dans le profil des utilisateurs et doit être unique : deux utilisateurs différents ne peuvent pas avoir le même identifiant.</p>"
        myTip+="<p>Si vous ne savez pas quelle valeur mettre ici, choisissez <b>stamboeknummer</b>, ce champ de SS correspond normalement au matricule des profs et des élèves.</p>"
        self.saisieChampCorrespondanceSS.setToolTip(myTip)
        mon_layout.addWidget(self.saisieChampCorrespondanceSS)
        
        #les boutons enregistrer, fermer, tester
        mon_bouton = QtWidgets.QPushButton('Enregistrer les valeurs et rafraichir la liste')
        mon_bouton.clicked.connect(self.recordValue)
        mon_layout.addWidget(mon_bouton)

        mon_bouton = QtWidgets.QPushButton('Fermer sans enregistrer')
        mon_bouton.clicked.connect(self.closeDialConfig)
        mon_layout.addWidget(mon_bouton)
        
        mon_bouton = QtWidgets.QPushButton('Tester la configuration')
        mon_bouton.clicked.connect(self.testConfig)
        mon_layout.addWidget(mon_bouton)

        # ajout du layout dans la fenetre
        self.dialConfig.setLayout(mon_layout)
        self.dialConfig.exec_()  # affichage de la boite
    #

    def closeDialConfig(self):
        self.dialConfig.close()
    
    def isConfigChanged(self):
        """
        Renvoie True si la config donnée dans la boite de dialogue est différente de celle stockée dans le dictConfig, sinon renvoie false
        """
        change=False
        if str(self.saisieUrl.text()) != self.dictConfig['urlEcole']:
            change=True
        if str(self.saisieApiKey.text()) != self.dictConfig['SSApiKey']:
            change=True
        if str(self.saisieInternummer.text()) != self.dictConfig['interNummerExpediteur']:
            change=True
        if str(self.saisieChampCorrespondanceSS.text()) != self.dictConfig['champIdentSS']:
            change=True
        return change
            

        
    def recordValue(self):
        """
        Enregistre les valeurs de configuration dans un fichier JSON uniquement s'il y a une différence entre le dictConfig et les valeurs données dans la boite de dialogue de configuration
        """
        if self.isConfigChanged() :
            TempDictConfig = {}
            TempDictConfig["urlEcole"] = str(self.saisieUrl.text())
            TempDictConfig["SSApiKey"] = str(self.saisieApiKey.text())
            TempDictConfig["interNummerExpediteur"] = str(self.saisieInternummer.text())
            TempDictConfig["champIdentSS"] = str(self.saisieChampCorrespondanceSS.text())
            
            myFile = open("config.json", "w")
            myFile.write(json.dumps(TempDictConfig))
            myFile.close()
            self.getConfig()  # apres avoir enregistré les valeurs de config dansle fichier JSON, on génère le dictConfig
            
         
        valid , message = self.isConfigValid()
        alert = "<div><p>Configuration non valide :<ul>"
        alert += message
        alert += "</ul></p></div>"

        if not valid :
            QtWidgets.QMessageBox.warning(self, 'Logiciel mal configuré', alert)
        else:
            self.prodDictUtilisateurs() #il faut recharger les utilisateurs au cas ou le champIdentSS a été placé sur une valeur différente de celle de départ
            self.dialConfig.close()

    def testConfig(self):
        """
        Teste la configuration en envoyant un message avec , si possible, une pièce jointe à la personne indiquée dans le champ 'expéditeur'.
        """
        valid , message = self.isConfigValid()
        if valid:
            try:
                self.client
            except:
                self.client=self.SSApiClient()
        
            
            try :
                myFile=open('./test.pdf', "rb")
                encodedFile = base64.b64encode(myFile.read())
                encodedFile = encodedFile.decode()  # on recupere la partie string du byte
                attachment = [{"filename": 'test.pdf', "filedata": encodedFile}]
                jsonAttachment = json.dumps(attachment)
                body="<p>Ceci est un message de test.</p>"
                body+="<p>Il contient un fichier pdf en pièce jointe.</p>"
                
            except FileNotFoundError:
                jsonAttachment=''
                body="<p>Ceci est un message de test.</p>"
                body+="<p>Il n'y a pas de fichier pdf en pièce jointe, le fichier 'test.pdf' n'est sans doute pas présent dans le répertoire où s'exécute le programme ayant envoyé ce emessage.</p>"
                
            fromaddr = self.dictConfig["interNummerExpediteur"]
            toaddr = self.dictConfig["interNummerExpediteur"]
            subject="Message de test"
            
            result = self.client.service.sendMsg(self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 0, 0)
            if result==0:
                QtWidgets.QMessageBox.information(self, 'Succès', "<p>Le message a bien été envoyé.</p>")
            else:
                message="<p>Une erreur est survenue.</p>"
                message+="<p>"+self.dictError[str(result)]+"</p>"
                QtWidgets.QMessageBox.warning(self, 'Echec', message)
        
        else : #la config n'est pas valide
            message="<p>Une erreur est survenue.</p>"
            message+="<p>"+self.dictError[str(result)]+"</p>"
            QtWidgets.QMessageBox.warning(self, 'Configuration non valide', message)
            
                
        

    ############################################
    #Fonctions concernant l'envoi des messages #
    #verifAvantEnvoi                           #
    #getAttachementName                        #
    #sendMsg                                   #
    #sendAllFiles                              #
    ############################################
    def verifAvantEnvoi(self):
        """Vérifie que tout est en ordre avant d'envoyer les fichiers"""
        valide , message = self.isConfigValid()
        alert = "<div><p>Vous voulez envoyer les messages mais :<ul>"
        alert += message
        
        # on vérifie s'il y a au moins un fichier à envoyer
        noFile = True
        for utilisateur in self.dictUtilisateurs.values():  
            if utilisateur.filePath != "":
                noFile = False
        if noFile:
            alert += "<li>Vous n'avez importé <b>aucun fichier</b> à envoyer</li>"
            valide = False
        
        #check if message subject is not empty
        if self.sujet.text() == "":  # or self.message.toHtml()=="":
            valide = False
            alert += "<li>Vous avez oublié le <b>sujet</b></li>"
        
        #check if message body is not empty
        if self.message.toPlainText() == "":
            valide = False
            alert += "<li>Vous avez oublié le <b>message</b></li>"
        
        #check if there is at least one recipient
        if (self.sendComptePrincipal.isChecked() == 0) & (self.sendComptesSecondaires.isChecked() == 0):
            valide = False
            alert += "<li>Vous n'avez sélectionné <b>aucun compte</b> pour l'envoi</li>"

        alert += "</ul></p></div>"

        if not valide :
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
    def getAttachementName(self , utilisateur):
        """Reçois un utilisateur comme argument et renvoit un nom formé à partir des données de l'utilisateur et respectant le motif donné dans les paramètres de renommage"""
        attachmentName=[]
        listMotif=self.renamePattern.text().split(';')
        for elem in listMotif:
            elem=elem.replace(' ','')
            if elem=='%nom':
                attachmentName.append(utilisateur.naam)
            elif elem=='%prenom':
                attachmentName.append(utilisateur.voornaam)
            else:
                attachmentName.append(elem)
        attachmentName.append('.pdf')
        attachmentName='_'.join(attachmentName)
        return(attachmentName)
                
                
    def sendMsg(self, utilisateur):
        """envoie le fichier fileName(str) à l'utilisateur utilisateur(obj) et aux comptes secondaires associés"""
        if self.Renommer.isChecked():
            attachmentName=self.getAttachementName(utilisateur)
        else:
            attachmentName=os.path.basename(utilisateur.filePath)
        #attachmentName = utilisateur.naam+"_"+utilisateur.voornaam+".pdf"
        with open(utilisateur.filePath, "rb") as myFile:
            encodedFile = base64.b64encode(myFile.read())
        encodedFile = encodedFile.decode()  # on recupere la partie string du byte
        attachment = [{"filename": attachmentName, "filedata": encodedFile}]

        jsonAttachment = json.dumps(attachment)

        fromaddr = self.dictConfig["interNummerExpediteur"]
        toaddr = utilisateur.internnummer
        subject = self.sujet.text()

        body = self.message.toHtml()
        if self.sendComptePrincipal.isChecked():
            result0 = self.client.service.sendMsg(
                self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 0, 0)
            if result0 == 0:
                utilisateur.statutMsgComptePrincipal = "Envoyé"
            else:
                utilisateur.statutMsgComptePrincipal = "Échec"
        else:
            result0 = 0  # s'il n'y a pas de message à envoyer au compte principal, il n'y a donc pas d'erreur liée à cet envoi

        if self.sendComptesSecondaires.isChecked():
            if utilisateur.status1 == "actief":
                result1 = self.client.service.sendMsg(
                    self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 1, 0)
                if result1 == 0:
                    utilisateur.statutMsgCoaccount1 = "Envoyé"
                else:
                    utilisateur.statutMsgCoaccount1 = "Échec"
            else:
                result1 = 0  # s'il n'y a pas de compte secondaire 1, il n'y a pas d'erreur liée à l'envoi sur ce compte
            if utilisateur.status2 == "actief":
                result2 = self.client.service.sendMsg(
                    self.dictConfig["SSApiKey"], toaddr, subject, body, fromaddr, jsonAttachment, 2, 0)
                if result2 == 0:
                    utilisateur.statutMsgCoaccount2 = "Envoyé"
                else:
                    utilisateur.statutMsgCoaccount2 = "Échec"
            else:
                result2 = 0  # s'il n'y a pas de compte secondaire 2, il n'y a pas d'erreur liée à l'envoi sur ce compte
        else:
            result1 = 0  # s'il n'y a pas de message à envoyer aux comptes secondaires, il n'y a donc pas d'erreur liée à cet envois
            result2 = 0  # idem

        # si aucun problème n'a été rencontré dans l'envoi des messages
        if (result0 == 0) & (result1 == 0) & (result2 == 0):
            # get the path to the file in the current directory
            # on récupère le path du fichier
            src = os.path.dirname(utilisateur.filePath)
            # rename the original file
            # on produit le nouveau nom du fichier
            newName = os.path.join(src, attachmentName)
            if self.Renommer.isChecked():
                os.rename(utilisateur.filePath, newName)  # on renomme le fichier si on a demandé de renommer
            utilisateur.filePath = ""  # on supprime le fichier à envoyer de l'objet Utilisateur

        self.updateFileList()
        QtWidgets.QApplication.processEvents()
    #

    def sendAllFiles(self):
        """Envoi tous les fichiers importés"""
        self.cleanDictUtilisateurs()
        try:
            self.client
        except:
            self.client=self.SSApiClient()
        
        for utilisateur in self.dictUtilisateurs.values():
            self.sendMsg(utilisateur)
        self.cleanDictUtilisateurs()
        if self.dictUtilisateurs == {}:
            message = "<div><p>Tous les documents ont été envoyés avec succès.</p><p>Si vous importez de nouveaux fichiers la liste des utilisateurs sera à nouveau importée.</p></div>"
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
        message += "<p>Ce programme renomme également les fichiers en leur donnant le nom et le prénom de l'utilisateur concerné par le fichier.</p><p> <b>Auteur : </b>J.N. Gautier</p>"
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
        message += "<p>Le nom des fichiers à envoyer doit correspondre au matricule de l'utilisateur, encodé dans le champ <b>stamboeknummer</b> de SmartSchool.</p>"
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


class Utilisateur():
    def __init__(self):
        self.voornaam = ""
        self.naam = ""
        self.internnummer = ""
        self.stamboeknummer = ""
        self.champIdentSS=""
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
    gui.prodDictUtilisateurs()
    gui.show()
    app.exec_()
