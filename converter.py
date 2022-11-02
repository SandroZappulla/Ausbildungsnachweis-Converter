# Python Ausbildungsnachweis PDF Converter
# Author: Sandro Zappulla
# version: v0.3

# IMPORT
# Installation:
#  python -m pip install --trusted-host files.pythonhosted.org --trusted-host pypi.org --trusted-host pypi.python.org docx2pdf
from docx2pdf import convert

# Ausbildungsnachweis Converter
print("Ausbildungsnachweis Converter von Sandro Zappulla")
print("-------------------------------------------------")
# Abfrage Counter
print("Ab Welcher Kalenderwoche soll die Konvertierung starten?")
kalender_woche_counter = int(input())
# Abfrage Max Wert
print("Bis zur Welcher Kalenderwoche soll Konvertiert werden?")
kalender_woche_max = int(input())

# Checks if the start value is smaller than the maximum value (if true: continue)
if kalender_woche_counter <= kalender_woche_max:
    # same condition to continue the loop for the convert
    while kalender_woche_counter <= kalender_woche_max:
        # definition filename/pdf_file (in/out put)
        filename = ("Ausbildungsnachweis KW" + str(kalender_woche_counter) + " SZ.docx")
        pdf_file = ("Ausbildungsnachweis KW" + str(kalender_woche_counter) + " SZ.pdf")

        # convert file from 'filename' to 'pdf_file'
        print("Erstelle PDF:", kalender_woche_counter, "/", kalender_woche_max)
        convert(filename, pdf_file)
    
        kalender_woche_counter+= 1
    print ("*** FERTIG ***")
else:
    print("FEHLER: Der Startwert darf nicht kleiner sein als der Maxwert.")
