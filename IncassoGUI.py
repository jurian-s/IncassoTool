# -*- coding: utf-8 -*-
"""
Created on Thu Apr 11 21:06:16 2024

@author: Jurian
"""

import sys
import IncassoTool
from PySide6 import QtWidgets, QtGui
from PySide6.QtCore import Qt


class IncassoGUI(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        
        self.LoadInputFileButton = QtWidgets.QPushButton("Laad input")
        self.LoadInputFileButton.clicked.connect(self.PushedLoadInputFileButton)
        
        self.Incasso = IncassoTool.IncassoTool()
        self.Incasso.LoadLedenBestand("Ledenbestand.json")
        self.Incasso.LoadFactuurnummers("Factuurnummers.json")
        
        self.LedenbestandDate = QtWidgets.QLabel(self.Incasso.LBCreationDate)
        
        self.EditFactuurnummersButton = QtWidgets.QPushButton("Activiteitencodes")
        self.EditFactuurnummersButton.clicked.connect(self.PushedEditFactuurnummersButton)
        
        self.ParseButton = QtWidgets.QPushButton("Genereer bestanden")
        self.ParseButton.clicked.connect(self.ParseandSave)
        
        self.ParseButton.setEnabled(False)
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.LoadInputFileButton)
        self.layout.addWidget(self.EditFactuurnummersButton)
        self.layout.addWidget(self.LedenbestandDate)
        self.layout.addWidget(self.ParseButton)
        
    def PushedLoadInputFileButton(self):
        self.fileDialog = QtWidgets.QFileDialog.getOpenFileName(self, "Open Document")
        try:
            self.Incasso.LoadInput(self.fileDialog[0])
            self.model = DataFrameModel(self.Incasso.InputFile)
            self.startTableView()
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
        
        self.Factuurlayout = QtWidgets.QVBoxLayout(self.FactuurnummerWidget)
        self.Factuurlayout.addWidget(self.factuur_view)
        self.Factuurlayout.addWidget(AddrowButton)
        self.Factuurlayout.addWidget(OKbutton)
        self.Factuurlayout.addWidget(Cancelbutton)
        
    def CloseFactuurnummerWidget(self):
        self.FactuurnummerWidget.close()
        
        
    def CloseandSaveFactuurnummerWidget(self):
        self.Incasso.FactuurDict = {}
        for row in range(0, self.Factuurmodel.rowCount()):
            Index0 = self.Factuurmodel.index(row, 0)
            Index1 = self.Factuurmodel.index(row, 1)
            Index2 = self.Factuurmodel.index(row, 2)
            Index3 = self.Factuurmodel.index(row, 3)
            self.Incasso.FactuurDict[self.Factuurmodel.data(Index0)] = {
                "DEBrekening": int(self.Factuurmodel.data(Index1)),
                "Tegenrekening": int(self.Factuurmodel.data(Index2)),
                "Kostenplaats": self.Factuurmodel.data(Index3)}
        self.Incasso.SaveFactuurnummers("Factuurnummers.json")
        self.FactuurnummerWidget.close()
        
    def ParseandSave(self):
        self.saveDialog = QtWidgets.QFileDialog.getExistingDirectory(self, "Save Directory")
        try:
            self.Incasso.ParseInput("01-01-2023")
        except ValueError as e:
            self.show_error_message(str(e))
        self.Incasso.SaveInput(f"{self.saveDialog}//Input_{self.Incasso.MsgId}.csv")
        self.Incasso.saveEBoekhoudenFactuur(f"{self.saveDialog}//eBoekhoudenFacturen_{self.Incasso.MsgId}.csv")
        self.Incasso.saveEBoekhoudenGiro(f"{self.saveDialog}//eBoekhoudenGiro_{self.Incasso.MsgId}.csv")
        self.Incasso.saveIncassoXML("emptyIncasso.xml", f"{self.saveDialog}//incassobatch_{self.Incasso.MsgId}.xml")
        self.Incasso.saveMailMerge(f"{self.saveDialog}//MailMerge_{self.Incasso.MsgId}.csv")
        
    def show_error_message(self, message):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setIcon(QtWidgets.QMessageBox.Warning)
        msg_box.setText("Error")
        msg_box.setInformativeText(message)
        msg_box.setWindowTitle("Error")
        msg_box.exec()    
        
        
    def AddRowToFactuur(self):
        self.Factuurmodel.appendRow([self.Factuurmodel.createItem(""), self.Factuurmodel.createItem(""), self.Factuurmodel.createItem(""), self.Factuurmodel.createItem("")])
        
        
        
    def startTableView(self):
        self.table_view = QtWidgets.QTableView()
        self.table_view.setModel(self.model)
        self.layout.addWidget(self.table_view)
        self.model.dataChanged.connect(self.updateIncassoInputFile)
        self.ColorTableWarnings()
        
    def ColorTableWarnings(self):
        self.Incasso.CheckNamen() 
        self.Incasso.CheckActCodes()
        for Ind in self.Incasso.InvalidNames:
            self.model.changeColor(Ind, 0, QtGui.QColorConstants.Red)
        for Ind in self.Incasso.PunchIbanNames:
            self.model.changeColor(Ind, 0, QtGui.QColorConstants.Yellow)
        for Ind in self.Incasso.InvalidCodes:
            self.model.changeColor(Ind, 5, QtGui.QColorConstants.Red)
        if len(self.Incasso.InvalidNames + self.Incasso.InvalidCodes) == 0:
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

    Main = IncassoGUI()
    Main.resize(800, 600)
    Main.show()

    sys.exit(app.exec())
    