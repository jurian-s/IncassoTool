# -*- coding: utf-8 -*-
"""
Created on Tue Apr  9 20:55:59 2024

@author: Jurian
"""

import pandas as pd
import json
import re
import logging
from datetime import datetime

class LedenbestandParser:
    def Parse(self, Path):
        File = pd.read_csv(Path, sep=";")
        Ledenbestand = File.to_dict(orient="index")
        self.IDdict = {}
        self.LedenDict = {}
        self.InvalidBIClist = []
        for Lid in Ledenbestand.values():
            if not self.is_valid_bic(Lid["Bank account BIC"]):
                logging.warning(f"BIC of {Lid['Name']} ({Lid['Bank account BIC']}) is not valid")
                self.InvalidBIClist.append(f"BIC of {Lid['Name']} ({Lid['Bank account BIC']}) is not valid")
            if not self.is_valid_iban(Lid["Bank account IBAN"]):
                raise ValueError(f"IBAN of {Lid['Name']} ({Lid['Bank account IBAN']})is not valid")
            self.IDdict[Lid['Name']] = Lid["Punch code"]
            self.LedenDict[Lid["Punch code"]] = {
                "Naam": Lid["Name"],
                "Voornaam": Lid["First name"],
                "Mail": Lid["Contact details email address"],
                "Incasso ID": Lid["Incasso id"],
                "Bank Holder Name": Lid["Bank account holder name"],
                "IBAN": Lid["Bank account IBAN"],
                "BIC": Lid["Bank account BIC"],
                "Mandate Date": Lid["Bank Mandate Date (dd-mm-yyyy)"]}
            
            
    def Save(self, Path):
        LedenBestand = {"Generation date": datetime.now().strftime("%d-%m-%Y"),
                        "ID": self.IDdict,
                        "LidByID": self.LedenDict}
        with open(Path, "w") as f:
            json.dump(LedenBestand, f, indent=4)
            
            
    def is_valid_bic(self, bic):
        pattern = re.compile(r'^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$')
        return bool(pattern.match(bic))
    
    def is_valid_iban(self, iban):
        pattern = re.compile(r'^[A-Z]{2}\d{2}[A-Z\d]{11,30}$')
        return bool(pattern.match(iban))
        
            
if __name__ == "__main__":
    Parser = LedenbestandParser()
    Parser.Parse(r"C:\Users\juria\Downloads\Ledenlijst alleen Jur 04-04-2024.csv")
    Parser.Save(r"C:\Users\juria\Downloads\Ledenbestand.json")