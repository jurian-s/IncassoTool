# -*- coding: utf-8 -*-
"""
Created on Thu Apr 11 21:06:16 2024

@author: Jurian
"""

import sys
import IncassoTool
import LedenbestandParser
from PySide6 import QtWidgets, QtGui
from PySide6.QtCore import Qt, QDate


class IncassoGUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        spacer = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.LoadInputFileButton = QtWidgets.QPushButton("Laad input")
        self.LoadInputFileButton.clicked.connect(self.PushedLoadInputFileButton)
        
        self.Incasso = IncassoTool.IncassoTool()
        self.Incasso.LoadLedenBestand("Ledenbestand.json")
        self.Incasso.LoadFactuurnummers("Factuurnummers.json")
        
        self.EditFactuurnummersButton = QtWidgets.QPushButton("Activiteitencodes")
        self.EditFactuurnummersButton.clicked.connect(self.PushedEditFactuurnummersButton)
        
        self.ParseButton = QtWidgets.QPushButton("Genereer bestanden")
        self.ParseButton.clicked.connect(self.ParseandSave)
        self.ParseButton.setEnabled(False)
        
        self.LedenBestandButton = QtWidgets.QPushButton("Update Ledenbestand")
        self.LedenBestandButton.clicked.connect(self.PushedRegenerateLedenbestandButton)
        
        self.date_edit = QtWidgets.QDateEdit()
        self.date_edit.setMinimumDate(QDate.currentDate().addDays(2))
        self.date_edit.setMaximumDate(QDate.currentDate().addMonths(3)) 
        self.date_edit.setDisplayFormat("dd-MM-yyyy")
        self.date_edit.setCalendarPopup(True)
        
        self.table_view = QtWidgets.QTableView()
        
        self.LedenbestandDate = QtWidgets.QLabel(f"Ledenbestand van: {self.Incasso.LBCreationDate}")
        self.IncassoDateLabel = QtWidgets.QLabel("Incasso uitvoeren op:")        
        ToolLabel = QtWidgets.QLabel("Incasso Tool")
        font = ToolLabel.font()
        font.setPointSize(20)
        font.setBold(True)
        font.setFamily("Proxima Nova")
        ToolLabel.setFont(font)
        pixmap = QtGui.QPixmap("punchlogo.png") 
        pixmap = pixmap.scaled(234, 50)
        Logo = QtWidgets.QLabel()
        Logo.setPixmap(pixmap)
        Logo.setAlignment(Qt.AlignRight)
        self.layout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.TopToBottom, self)
        
        HeaderLayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.LeftToRight) 
        HeaderLayout.addWidget(ToolLabel)
        HeaderLayout.addWidget(Logo)
        
        self.layout.addLayout(HeaderLayout, 1)
        self.layout.addWidget(self.LoadInputFileButton)     
        self.layout.addWidget(self.table_view, 8)
        
        LedenBestandLayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.TopToBottom)
        LedenBestandLayout.addSpacerItem(spacer)
        LedenBestandLayout.addWidget(self.LedenbestandDate)
        LedenBestandLayout.addWidget(self.LedenBestandButton)
        IncassoOptiesLayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.TopToBottom)
        IncassoOptiesLayout.addWidget(self.IncassoDateLabel)
        IncassoOptiesLayout.addWidget(self.date_edit)
        IncassoOptiesLayout.addWidget(self.ParseButton)
        
        FactuurOptiesLayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.TopToBottom)
        FactuurOptiesLayout.addSpacerItem(spacer)
        FactuurOptiesLayout.addWidget(self.EditFactuurnummersButton)
        
        OptieBalkLayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.LeftToRight)
        OptieBalkLayout.addLayout(LedenBestandLayout)
        OptieBalkLayout.addLayout(FactuurOptiesLayout)
        OptieBalkLayout.addLayout(IncassoOptiesLayout)
        self.layout.addLayout(OptieBalkLayout, 2)
        
    
        
    def PushedRegenerateLedenbestandButton(self):
        self.LedenbestandDialog = QtWidgets.QFileDialog.getOpenFileName(self, "Open Ledenbestand", filter="CSV file (*.csv)")
        Parser = LedenbestandParser.LedenbestandParser()
        try:
            Parser.Parse(self.LedenbestandDialog[0])
            Parser.Save("Ledenbestand.json")
            self.Incasso.LoadLedenBestand("Ledenbestand.json")
            self.LedenbestandDate.setText(f"Ledenbestand van: {self.Incasso.LBCreationDate}")
            if len(Parser.InvalidBIClist) > 0:
                self.show_BICwarning(Parser.InvalidBIClist)
        except ValueError as e:
            self.show_error_message(str(e))
        
        
            
    def PushedLoadInputFileButton(self):
        self.fileDialog = QtWidgets.QFileDialog.getOpenFileName(self, "Open Input Document", filter="Excel files (*.csv  *.xls *.xlsx *.xlsm)")
        try:
            self.Incasso.LoadInput(self.fileDialog[0])
            self.model = DataFrameModel(self.Incasso.InputFile)
            self.table_view.setModel(self.model)
            self.model.dataChanged.connect(self.handleDataChanged)
            self.ColorTableWarnings()
        except ValueError as e:
            self.show_error_message(str(e))
            
        
    def PushedEditFactuurnummersButton(self):
        self.FactuurnummerWidget = QtWidgets.QWidget()
        self.FactuurnummerWidget.setWindowTitle("Activiteitencodes")
        self.FactuurnummerWidget.show()
        
        self.Factuurmodel = FactuurDictModel(self.Incasso.FactuurDict)
        self.factuur_view = QtWidgets.QTableView()
        self.factuur_view.setModel(self.Factuurmodel)
        
        
        AddrowButton = QtWidgets.QPushButton("Voeg rij toe")
        AddrowButton.clicked.connect(self.AddRowToFactuur)
        
        Cancelbutton = QtWidgets.QPushButton("Annuleren")
        Cancelbutton.clicked.connect(self.CloseFactuurnummerWidget)
        
        OKbutton = QtWidgets.QPushButton("Opslaan")
        OKbutton.clicked.connect(self.CloseandSaveFactuurnummerWidget)
        
        self.Factuurlayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.TopToBottom, self.FactuurnummerWidget)
        self.Factuurlayout.addWidget(self.factuur_view)
        self.Factuurlayout.addWidget(AddrowButton)
        CloseWindowLayout = QtWidgets.QBoxLayout(QtWidgets.QBoxLayout.LeftToRight)
        CloseWindowLayout.addWidget(OKbutton)
        CloseWindowLayout.addWidget(Cancelbutton)
        self.Factuurlayout.addLayout(CloseWindowLayout)
        
    def CloseFactuurnummerWidget(self):
        self.FactuurnummerWidget.close()
        
        
    def CloseandSaveFactuurnummerWidget(self):
        self.Incasso.FactuurDict = {}
        for row in range(0, self.Factuurmodel.rowCount()):
            Index0 = self.Factuurmodel.index(row, 0)
            Index1 = self.Factuurmodel.index(row, 1)
            Index2 = self.Factuurmodel.index(row, 2)
            Index3 = self.Factuurmodel.index(row, 3)
            if self.Factuurmodel.data(Index0) != "":
                self.Incasso.FactuurDict[self.Factuurmodel.data(Index0)] = {
                    "DEBrekening": int(self.Factuurmodel.data(Index1)),
                    "Tegenrekening": int(self.Factuurmodel.data(Index2)),
                    "Kostenplaats": self.Factuurmodel.data(Index3)}
        self.Incasso.SaveFactuurnummers("Factuurnummers.json")
        self.FactuurnummerWidget.close()
        
    def ParseandSave(self):
        self.saveDialog = QtWidgets.QFileDialog.getExistingDirectory(self, "Save Directory")
        try:
            print(self.date_edit.date().toString("yyyy-MM-dd"))
            self.Incasso.ParseInput(self.date_edit.date().toString("yyyy-MM-dd"))
            self.Incasso.SaveInput(f"{self.saveDialog}//Input_{self.Incasso.MsgId}.csv")
            self.Incasso.saveEBoekhoudenFactuur(f"{self.saveDialog}//eBoekhoudenFacturen_{self.Incasso.MsgId}.csv")
            self.Incasso.saveEBoekhoudenGiro(f"{self.saveDialog}//eBoekhoudenGiro_{self.Incasso.MsgId}.csv")
            self.Incasso.saveIncassoXML("emptyIncasso.xml", f"{self.saveDialog}//incassobatch_{self.Incasso.MsgId}.xml")
            self.Incasso.saveMailMerge(f"{self.saveDialog}//MailMerge_{self.Incasso.MsgId}.csv")
            msg_box = QtWidgets.QMessageBox()
            msg_box.setText("Succes!")
            msg_box.setInformativeText(f"De incassobestanden zijn succesvol opgeslagen in de volgende locatie: {self.saveDialog}")
            msg_box.setWindowTitle("Incasso opgeslagen")
            msg_box.exec() 
        except ValueError as e:
            self.show_error_message(str(e))
        
        
        
    def show_error_message(self, message):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setText("Error")
        msg_box.setInformativeText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec()    
        
    def show_BICwarning(self, warninglist):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setText("Warning")
        msg_box.setInformativeText("\n".join(warninglist))
        msg_box.setWindowTitle("Warning")
        msg_box.exec()    
        
        
    def AddRowToFactuur(self):
        self.Factuurmodel.appendRow([self.Factuurmodel.createItem(""), self.Factuurmodel.createItem(""), self.Factuurmodel.createItem(""), self.Factuurmodel.createItem("")])
        
    def handleDataChanged(self, topLeft, bottomRight, roles=None):
        if roles is None or Qt.DisplayRole in roles:
            self.updateIncassoInputFile(topLeft, bottomRight)

    def ColorTableWarnings(self):
        self.Incasso.CheckNamen() 
        self.Incasso.CheckActCodes()
        self.model.resetColor()
        for Ind in self.Incasso.InvalidNames:
            self.model.changeColor(Ind, 0, QtGui.QColorConstants.Red)
        for Ind in self.Incasso.PunchIbanNames:
            self.model.changeColor(Ind, 0, QtGui.QColorConstants.Yellow)
        for Ind in self.Incasso.InvalidCodes:
            self.model.changeColor(Ind, 5, QtGui.QColorConstants.Red)
        for Ind in self.Incasso.InvalidBICs:
            self.model.changeColor(Ind, 0, QtGui.QColorConstants.DarkYellow)
        if len(self.Incasso.InvalidNames) + len(self.Incasso.InvalidCodes) == 0:
            self.ParseButton.setEnabled(True)
        else:
            self.ParseButton.setEnabled(False)
            
            
                    

    def updateIncassoInputFile(self, topLeft, bottomRight):
        # Extract changed data and update Incasso.InputFile
        for index in self.table_view.selectionModel().selectedIndexes():
            row = index.row()
            col = index.column()
            value = self.model.data(index)
            if self.Incasso.InputFile.dtypes.iloc[col] == 'float64':
                 value = float(value)
            elif self.Incasso.InputFile.dtypes.iloc[col] == 'int64':
                 value = int(value)
            self.Incasso.InputFile.iloc[row, col] = value
        self.ColorTableWarnings()



class DataFrameModel(QtGui.QStandardItemModel):
    def __init__(self, dataframe):
        super().__init__(dataframe.shape[0], dataframe.shape[1])
        self.dataframe = dataframe

        self.setHorizontalHeaderLabels(dataframe.columns)

        for i in range(dataframe.shape[0]):
            for j in range(dataframe.shape[1]):
                item = str(dataframe.iloc[i, j])
                self.setItem(i, j, self.createItem(item))

    def createItem(self, text):
        item = QtGui.QStandardItem(text)
        item.setEditable(True)
        return item
    def resetColor(self):
        for row in range(self.rowCount()):
            for col in range(self.columnCount()):
                self.item(row, col).setBackground(QtGui.QColor(Qt.transparent))
    def changeColor(self, row, col, color):
        self.item(row, col).setBackground(color)
    
class FactuurDictModel(QtGui.QStandardItemModel):
    def __init__(self, Dictionary):
        super().__init__(len(Dictionary.keys()), 3)
        self.dict = Dictionary
        self.setHorizontalHeaderLabels(["Activiteitencode", "DebRekening", "Tegenrekening", "Kostenplaats"])
        for i, key in enumerate(Dictionary.keys()):
            self.setItem(i, 0, self.createItem(key))
            self.setItem(i, 1, self.createItem(str(Dictionary[key]["DEBrekening"])))
            self.setItem(i, 2, self.createItem(str(Dictionary[key]["Tegenrekening"])))
            self.setItem(i, 3, self.createItem(Dictionary[key]["Kostenplaats"]))
    def createItem(self, text):
        item = QtGui.QStandardItem(text)
        item.setEditable(True)
        return item
        
        
        
        
if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    app.setWindowIcon(QtGui.QIcon("PunchIcon.ico"))
    app.setApplicationDisplayName("Incasso Tool")
    Main = IncassoGUI()
    Main.resize(800, 600)
    
    Main.show()

    sys.exit(app.exec())
    