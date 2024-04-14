#  How to Build
Use python 3.11 to make a virtual environment
Install required modules with "pip install requirements.txt"
Build the .exe with "pyside6-deploy IncassoGUI.py"
(If you don't want the console to be visible, you can use the generated spec file, add --disable-console in the extra-args for Nuitka)

#  How to use
Start IncassoGUI.exe, en laad de input.xls
Zorg ervoor dat het ledenbestand up to date is, je kan deze updaten door op Update Ledenbestand te drukken, en een ledenbestand.csv te selecteren die je hebt gedownload van de website. 

Om Activiteitcodes toe te voegen of aan te passen, druk op de Activiteitencodes knop. 

Als een naam rood wordt, staat deze naam niet in het ledenbestand (Waarschijnlijk een typefout), Je kan de naam in de tool zelf aanpassen. 
Als een naam lichtgeel wordt, dan is de IBAN die bij deze naam hoort volgens het Ledenbestand gelijk aan die van Punch. Deze transactie zal wel een factuur aanmaken, maar niet geïncasseerd worden (En dus ook niet in E-boekhouden Giro bestand komen, en niet in de mail merge). Bij donkergeel is dit hetzelfde, alleen komt het doordat de BIC niet een valid vorm heeft. 

Als een activiteitcode rood is, dan mist deze in de activiteitencodeslijst. 

De incasso kan pas gegenereerd worden als er geen rode namen en activiteitencodes meer zijn. Vergeet niet de incassodatum aan te passen!

Als je op Genereer bestanden drukt, kan je de directory kiezen waar de bestanden worden opgeslagen. 
