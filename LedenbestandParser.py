# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 20:55:59 2024

@author: Jurian
"""

import pandas as pd
import json
from datetime import datetime

class LedenbestandParser:
    def Parse(self, Path):
        File = pd.read_csv(Path, sep=";")
        Ledenbestand = File.to_dict(orient="index")
        self.IDdict = {}
        self.LedenDict = {}
        for Lid in Ledenbestand.values():
            self.IDdict[Lid['Name']] = Lid["Punch code"]
            self.LedenDict[Lid["Punch code"]] = {
                "Naam": Lid["Name"],
                "Voornaam": Lid["First name"],
                "Mail": Lid["Contact details email address"],
                "Bank Holder name": Lid["Bank account holder name"],
                "IBAN": Lid["Bank account IBAN"],
                "BIC": Lid["Bank account BIC"],
                "Mandate date": Lid["Bank Mandate Date (dd-mm-yyyy)"]}
            
            
    def Save(self, Path):
        LedenBestand = {"Generation date": datetime.now().strftime("%d-%m-%Y"),
                        "ID": self.IDdict,
                        "LidByID": self.LedenDict}
        with open(Path, "w") as f:
            json.dump(LedenBestand, f, indent=4)
        
            
if __name__ == "__main__":
    Parser = LedenbestandParser()
    Parser.Parse(r"C:\Users\juria\Downloads\Ledenlijst alleen Jur 04-04-2024.csv")
    Parser.Save(r"C:\Users\juria\Downloads\Ledenbestand.json")