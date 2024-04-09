# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 21:17:59 2024

@author: Jurian
"""


import pandas as pd
import json
from datetime import datetime
import xmltodict





class IncassoTool: 
    def LoadLedenBestand(self, path):
        with open(path) as f:
            Ledenbestand = json.load(f)
        self.IDdict = Ledenbestand["ID"]
        self.Leden = Ledenbestand["LidByID"]
        
    def LoadFactuurnummers(self, path):
        File = pd.read_excel(ImportPath, sheet_name="E-boekhoud Factuurnummers")
        self.FactuurDict = {}
        Filedict = File.to_dict(orient="index")
        for row in Filedict.values():
            self.FactuurDict[row["Activiteitscode"]] = {
                "DEBrekening": row["DEBrekening"],
                "Tegenrekening": row["Tegenrekening"],
                'Kostenplaats': row["Kostenplaats"]}
        
        
    
    def ParseInput(self, ImportPath, IncassoDate):
        File = pd.read_excel(ImportPath, sheet_name="Invoer", skiprows=4)
        InputDict = File.to_dict(orient="index")
        CurrentDatetime = datetime.now()
        self.IncassoDate = IncassoDate
        self.IncassoNaam = CurrentDatetime.strftime("Incasso DSVV Punch %B '%y")
        self.MsgId = CurrentDatetime.strftime("PUNDD%Y%m%d%H%M%S")
        self.CreationTime = CurrentDatetime.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]
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
        self.nTxs = 0
        for transaction in InputDict.values():
            self.transaction = transaction
            ID = self.IDdict[transaction['Naam']]
            if ID not in self.TransactionDict.keys():
                self.addIDtoTransactionDict(ID)
                self.nTxs = self.nTxs + 1
            self.TransSum = self.TransSum + transaction["Bedrag"]
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
        if str(Trans["Factuurnummer"]) != "nan":
            self.EboekhoudenGiro["Factuurnummer"].append(Trans["Factuurnummer"])
        else:
            self.EboekhoudenGiro["Factuurnummer"].append("{ID}-{ActCode}-{MY}".format(ID = ID[2:].replace("-", ""), ActCode=Trans['Activiteitcode'], MY =MY))
        self.EboekhoudenGiro["Tegenrekening"].append(self.FactuurDict[Trans["Activiteitcode"]]["DEBrekening"])
        self.EboekhoudenGiro["Rekening"].append("11910")
        self.EboekhoudenGiro["Boekstuk"].append("GI" + MY)
        self.Trans = Trans
        if str(Trans["Factuurnummer"]) == "nan":
            self.EboekhoudenFact["Datum"].append(Trans["Datum van activiteit"])
            self.EboekhoudenFact["Omschrijving"].append("{oms} - {naam}".format(oms = Trans["Omschrijving Nederlands"], naam= Trans["Naam"]))
            self.EboekhoudenFact["Bedrag (EUR)"].append(Trans["Bedrag"])
            self.EboekhoudenFact["MutatieSoort"].append("Factuur verstuurd")
            self.EboekhoudenFact["Relatie"].append("P"+ID)        
            self.EboekhoudenFact["Factuurnummer"].append("{ID}-{ActCode}-{MY}".format(ID = ID[2:].replace("-", ""), ActCode=Trans['Activiteitcode'], MY =MY))
            self.EboekhoudenFact["Tegenrekening"].append(self.FactuurDict[Trans["Activiteitcode"]]["Tegenrekening"])
            self.EboekhoudenFact["Rekening"].append(self.FactuurDict[Trans["Activiteitcode"]]["DEBrekening"])
            self.EboekhoudenFact["Boekstuk"].append("GI" + MY)
            self.EboekhoudenFact["Betalingstermijn"].append(30)
            self.EboekhoudenFact["Kostenplaats"].append(self.FactuurDict[Trans["Activiteitcode"]]["Kostenplaats"])
            
            
    def saveEBoekhoudenGiro(self, Path):
        GiroDF = pd.DataFrame(self.EboekhoudenGiro)
        GiroDF.to_csv(Path)
        
    def saveEBoekhoudenFactuur(self, Path):
        FactuurDF = pd.DataFrame(self.EboekhoudenFact)
        FactuurDF.to_csv(Path)
        
    def saveIncassoXML(self, EmptyXMLPath, Path):
        with open(EmptyXMLPath) as file:
                XML = xmltodict.parse(file.read())
        XML["Document"]["CstmrDrctDbtInitn"]["GrpHdr"]["MsgId"] = self.MsgId
        XML["Document"]["CstmrDrctDbtInitn"]["GrpHdr"]["CreDtTm"] = self.CreationTime
        XML["Document"]["CstmrDrctDbtInitn"]["GrpHdr"]["NbOfTxs"] = str(self.nTxs)
        XML["Document"]["CstmrDrctDbtInitn"]["GrpHdr"]["CtrlSum"] = str(self.TransSum)
        XML["Document"]["CstmrDrctDbtInitn"]["PmtInf"]["PmtInfID"] = self.MsgId + "A"
        XML["Document"]["CstmrDrctDbtInitn"]["PmtInf"]["NbOfTxs"] = str(self.nTxs)
        XML["Document"]["CstmrDrctDbtInitn"]["PmtInf"]["CtrlSum"] = str(self.TransSum)
        XML["Document"]["CstmrDrctDbtInitn"]["PmtInf"]["ReqdColltnDt"] = self.IncassoDate

        DDTransactions = [{}] * self.nTxs
        for Ind, ID in enumerate(self.TransactionDict):
            LidInfo = self.Leden[ID]
            DDTransactions[Ind] = {'PmtId': {'EndToEndId': f"{self.MsgId}A-{Ind:>04}"},
                                   'InstdAmt': {'@Ccy': 'EUR', '#text': str(self.TransactionDict[ID]["TransactionSum"])},
                                   'DrctDbtTx': {'MndtRltdInf': {'MndtId': LidInfo["Incasso ID"],
                                   'DtOfSgntr': LidInfo["Mandate Date"]}},
                                   'DbtrAgt': {'FinInstnId': {'BIC': LidInfo["BIC"]}},
                                   'Dbtr': {'Nm': LidInfo["Bank Holder Name"]},
                                   'DbtrAcct': {'Id': {'IBAN': LidInfo["IBAN"]}},
                                   'RmtInf': {'Ustrd': self.IncassoNaam}}
        XML["Document"]["CstmrDrctDbtInitn"]["PmtInf"]["DrctDbtTxInf"] = DDTransactions
        with open(Path, "w") as file:
           file.write(xmltodict.unparse(XML, pretty=True))
            
        
        
    
if __name__ == "__main__":
    Tool = IncassoTool()
    Tool.LoadLedenBestand(r"C:\Users\juria\Downloads\Ledenbestand.json")
    ImportPath = r"C:\Users\juria\Downloads\Incassojetser Feb goed 23.xlsm"
    Tool.LoadFactuurnummers(ImportPath)
    
    IncassoDate = "23-04-2024"
    Tool.ParseInput(ImportPath, IncassoDate)
    Tool.saveEBoekhoudenGiro(r"C:\Users\juria\Downloads\FebEboekhoudenGiro.csv")
    Tool.saveEBoekhoudenFactuur(r"C:\Users\juria\Downloads\FebEboekhoudenFact.csv")
    Tool.saveIncassoXML(r"C:\Users\juria\Downloads\PennieTools\emptyIncasso.xml", r"C:\Users\juria\Downloads\PennieTools\Incasso.xml")

        

        
            