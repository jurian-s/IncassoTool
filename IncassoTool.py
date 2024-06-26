# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 21:17:59 2024

@author: Jurian
"""


import pandas as pd
import json
from datetime import datetime
import xmltodict
import re




class IncassoTool: 
    def LoadLedenBestand(self, path):
        with open(path) as f:
            Ledenbestand = json.load(f)
        self.LBCreationDate = Ledenbestand["Generation date"]
        self.IDdict = Ledenbestand["ID"]
        self.Leden = Ledenbestand["LidByID"]
        
    def LoadFactuurnummers(self, path):
        with open(path) as f:
            self.FactuurDict = json.load(f)
            
    def LoadFactuurnummersFromExcel(self, path):
        File = pd.read_excel(path, sheet_name="E-boekhoud Factuurnummers")
        self.FactuurDict = {}
        Filedict = File.to_dict(orient="index")
        for row in Filedict.values():
            self.FactuurDict[row["Activiteitscode"]] = {
                "DEBrekening": row["DEBrekening"],
                "Tegenrekening": row["TegenreFackening"],
                'Kostenplaats': row["Kostenplaats"]}
        
    def SaveFactuurnummers(self, path):
        with open(path, "w") as f:
            json.dump(self.FactuurDict, f, indent=4)   
    
    def LoadInput(self, ImportPath):
        if ImportPath.split('.')[-1] == "csv":
            self.InputFile = pd.read_csv(ImportPath)            
        else:            
            self.InputFile = pd.read_excel(ImportPath, sheet_name="Invoer", skiprows=1)
        if self.InputFile.keys().tolist() != ['Naam','Bedrag','Omschrijving Nederlands','Omschrijving Engels','Datum van activiteit','Activiteitcode','Factuurnummer']:
            raise ValueError("Input bestand bevat niet de correcte headers, controleer of je het juiste bestand importeert")
            
    def SaveInput(self, ImportPath):
        self.InputFile.to_csv(ImportPath, index=False)
        
    def CheckActCodes(self):
        self.InvalidCodes = []
        for Ind, Code in enumerate(self.InputFile["Activiteitcode"]):
            if Code not in self.FactuurDict.keys():
                self.InvalidCodes.append(Ind)
    
    def CheckNamen(self):
        self.InvalidNames = []
        self.PunchIbanNames = []
        self.InvalidBICs = []
        for Ind, Naam in enumerate(self.InputFile["Naam"]):
            if Naam not in self.IDdict.keys():
                self.InvalidNames.append(Ind)
            elif not self.is_valid_bic(self.Leden[self.IDdict[Naam]]["BIC"]):
                self.InvalidBICs.append(Ind)
            elif self.Leden[self.IDdict[Naam]]["IBAN"] == "NL33SNSB0339513241":
                self.PunchIbanNames.append(Ind)
        
    def ParseInput(self, IncassoDate):
        InputDict = self.InputFile.to_dict(orient="index")
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
        self.maxTypes = 0
        for transaction in InputDict.values():
            ID = self.IDdict[transaction['Naam']]
            Incasseerbaar = True
            if not self.is_valid_bic(self.Leden[ID]["BIC"]) or self.Leden[ID]["IBAN"] == "NL33SNSB0339513241":
                Incasseerbaar = False
            else:
                if ID not in self.TransactionDict.keys():
                    self.addIDtoTransactionDict(ID)
                    self.nTxs = self.nTxs + 1
                self.TransSum = self.TransSum + transaction["Bedrag"]
                self.TransactionDict[ID]["Oms"].append(transaction['Omschrijving Nederlands'])
                self.TransactionDict[ID]["OmsEng"].append(transaction["Omschrijving Engels"])  
                self.TransactionDict[ID]["Bedrag"].append(transaction["Bedrag"])  
                self.maxTypes = max(self.maxTypes, len(self.TransactionDict[ID]['Oms']))
                self.TransactionDict[ID]["TransactionSum"] = self.TransactionDict[ID]["TransactionSum"] + transaction["Bedrag"]
            self.addTransToEboekhouden(transaction, ID, Incasseerbaar)
            
    
     
            
    def addIDtoTransactionDict(self, ID):
        self.TransactionDict[ID] = {}
        self.TransactionDict[ID]["Oms"] = []
        self.TransactionDict[ID]["OmsEng"] = []
        self.TransactionDict[ID]["Bedrag"] = []
        self.TransactionDict[ID]["TransactionSum"] = 0
        self.TransactionDict[ID]["TransactionBoekHouden"] = []
        
    def addTransToEboekhouden(self, Trans, ID, Incasseerbaar):
        date = datetime.strptime(self.IncassoDate, '%Y-%m-%d')
        MY = date.strftime("%y{}").format(date.month)
        if Incasseerbaar:
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
            
    def is_valid_bic(self, bic):
        pattern = re.compile(r'^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$')
        return bool(pattern.match(bic))
    
    
    def saveEBoekhoudenGiro(self, Path):
        GiroDF = pd.DataFrame(self.EboekhoudenGiro)
        GiroDF.to_csv(Path, index=False)
        
    def saveEBoekhoudenFactuur(self, Path):
        FactuurDF = pd.DataFrame(self.EboekhoudenFact)
        FactuurDF.to_csv(Path, index=False)
        
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
            Lid = self.Leden[ID]
            DDTransactions[Ind] = {'PmtId': {'EndToEndId': f"{self.MsgId}A-{Ind:>04}"},
                                   'InstdAmt': {'@Ccy': 'EUR', '#text': str(self.TransactionDict[ID]["TransactionSum"])},
                                   'DrctDbtTx': {'MndtRltdInf': {'MndtId': Lid["Incasso ID"],
                                   'DtOfSgntr': Lid["Mandate Date"]}},
                                   'DbtrAgt': {'FinInstnId': {'BIC': Lid["BIC"]}},
                                   'Dbtr': {'Nm': Lid["Bank Holder Name"]},
                                   'DbtrAcct': {'Id': {'IBAN': Lid["IBAN"]}},
                                   'RmtInf': {'Ustrd': self.IncassoNaam}}
        XML["Document"]["CstmrDrctDbtInitn"]["PmtInf"]["DrctDbtTxInf"] = DDTransactions
        with open(Path, "w") as file:
           file.write(xmltodict.unparse(XML, pretty=True))
           
    def saveMailMerge(self, Path):
        MailMergeDict = {}
        for ID in self.TransactionDict:
            Trans = self.TransactionDict[ID]
            Lid = self.Leden[ID]
            MailMergeDict[ID] = [Lid["Voornaam"], Lid["Mail"], Trans["TransactionSum"]]
            for Ind, Oms in enumerate(Trans["Oms"]):
                MailMergeDict[ID].append(Oms)
                MailMergeDict[ID].append(Trans["OmsEng"][Ind])
                MailMergeDict[ID].append(Trans["Bedrag"][Ind])
        MailMergeHeader = ["Naam", "Email", "Totaal"]
        for i in range(self.maxTypes):
            MailMergeHeader.append(f"Oms{i+1}a")
            MailMergeHeader.append(f"Oms{i+1}b")
            MailMergeHeader.append(f"Bedrag{i+1}")
        MailMergeDF = pd.DataFrame.from_dict(MailMergeDict, orient="index", columns=MailMergeHeader)
        MailMergeDF.to_csv(Path, index=False)
            
            
        
        
    
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
    Tool.saveMailMerge(r"C:\Users\juria\Downloads\MailMerge.csv")

        

        
            