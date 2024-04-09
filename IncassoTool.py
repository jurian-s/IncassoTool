# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 21:17:59 2024

@author: Jurian
"""


import pandas as pd
import json
from datetime import datetime





class IncassoTool: 
    def LoadLedenBestand(self, path):
        with open(path) as f:
            Ledenbestand = json.load(f)
        self.IDdict = Ledenbestand["ID"]
        self.Leden = Ledenbetandt["LidByID"]
        
    def LoadFactuurnummers(self, path):
        with open(path) as f:
            self.FactuurDict = json.load(f)
    
    def ParseInput(self, ImportPath, IncassoDate):
        File = pd.read_excel(ImportPath, sheet_name="Invoer", skiprows=5)
        InputDict = File.to_dict(orient="index")
        self.IncassoDate = IncassoDate
        self.TransactionDict = {}
        self.EboekhoudenGiro = {"Datum": [],
                                "Omschrijving": [],
                                "Bedrag (EUR)": [],
                                "MutatieSoort": [],
                                "Relatie": [],
                                "Factuurnummer": [],
                                "Tegenrekening": [],
                                "Rekening": [],
                                "Boekstuk": []}
        self.EboekhoudenFact = {"Datum": [],
                                "Omschrijving": [],
                                "Bedrag (EUR)": [],
                                "MutatieSoort": [],
                                "Relatie": [],
                                "Factuurnummer": [],
                                "Tegenrekening": [],
                                "Rekening": [],
                                "Betalingstermijn": [],
                                "Kostenplaats": [],
                                "Boekstuk": []}
        self.TransSum = 0
        for transaction in InputDict:
            ID = self.IDdict[transaction['Naam']]
            if ID not in self.TransactionDict.keys():
                self.addIDtoTransactionDict(ID)
            self.TransSum =+ transaction["Bedrag"]
            self.TransactionDict[ID]["Oms"].append(transaction['Omschrijving Nederlands'])
            self.TransactionDict[ID]["OmsEng"].append(transaction["Omschrijving Engels"])        
            self.TransactionDict[ID]["TransactionSum"] = self.TransactionDict[ID]["TransactionSum"] + transaction["Bedrag"]
            self.addTransToEboekhouden(transaction, ID)
            
    
     
            
    def addIDtoTransactionDict(self, ID):
        self.TransactionDict[ID] = {}
        self.TransactionDict[ID]["Oms"] = []
        self.TransactionDict[ID]["OmsEng"] = []
        self.TransactionDict[ID]["TransactionSum"] = 0
        self.TransactionDict[ID]["TransactionBoekHouden"] = []
        
    def addTransToEboekhouden(self, Trans, ID):
        date = datetime.strptime(self.IncassoDate, '%d-%m-%Y')
        MY = date.strftime("%y{}").format(date.month)
        self.EboekhoudenGiro["Datum"].append(self.IncassoDate)
        self.EboekhoudenGiro["Omschrijving"].append("{oms} - {naam}".format(oms = Trans["Omschrijving Nederlands"], naam= Trans["Naam"]))
        self.EboekhoudenGiro["Bedrag (EUR)"].append(Trans["Bedrag"])
        self.EboekhoudenGiro["MutatieSoort"].append("Factuurbetaling ontvangen")
        self.EboekhoudenGiro["Relatie"].append("P"+ID)
        if Trans["Factuurnummer"] != "nan":
            self.EboekhoudenGiro["Factuurnummer"].append(Trans["Factuurnummer"])
        else:
            self.EboekhoudenGiro["Factuurnummer"].append("{ID}-{ActCode}-{MY}".format(ID = ID[2:].replace("-", ""), ActCode=Trans['Activiteitcode'], MY =MY))
        self.EboekhoudenGiro["Tegenrekening"].append(self.FactuurDict[Trans["Activiteitcode"]["DEBrekening"]])
        self.EboekhoudenGiro["Rekening"].append("11910")
        self.EboekhoudenGiro["Boekstuk"].append("GI" + MY)

        if Trans["Factuurnummer"] == "nan":
            self.EboekhoudenFact["Datum"].append(Trans["Datum van Activiteit"])
            self.EboekhoudenFact["Omschrijving"].append("{oms} - {naam}".format(oms = Trans["Omschrijving Nederlands"], naam= Trans["Naam"]))
            self.EboekhoudenFact["Bedrag (EUR)"].append(Trans["Bedrag"])
            self.EboekhoudenFact["MutatieSoort"].append("Factuur verstuurd")
            self.EboekhoudenFact["Relatie"].append("P"+ID)        
            self.EboekhoudenFact["Factuurnummer"].append("{ID}-{ActCode}-{MY}".format(ID = ID[2:].replace("-", ""), ActCode=Trans['Activiteitcode'], MY =MY))
            self.EboekhoudenFact["Tegenrekening"].append(self.FactuurDict[Trans["Activiteitcode"]["Tegenrekening"]])
            self.EboekhoudenFact["Rekening"].append(self.FactuurDict[Trans["Activiteitcode"]["DEBrekening"]])
            self.EboekhoudenFact["Boekstuk"].append("GI" + MY)
            self.EboekhoudenFact["Betalingstermijn"].append(30)
            self.EboekhoudenFact["Kostenplaats"].append(self.FactuurDict[Trans["Activiteitcode"]["Kostenplaats"]])
            
            
            
if __name__ == "__main__":
    Tool = IncassoTool()
    Tool.LoadLedenBestand()
    Tool.LoadFactuurnummers()
    Tool.ParseInput(ImportPath, IncassoDate)

        

        
            