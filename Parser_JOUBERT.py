#-------------------------------------------------------------------------------
# Name:        module2
# Purpose:
#
# Author:      Benjamin
#
# Created:     24/09/2020
# Copyright:   (c) Benjamin 2020
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os
import datetime
import encodings
import requests
import xlwings as xw
from bs4 import BeautifulSoup

aujourdhui =  datetime.date.today()
aujourdhui = str(aujourdhui)

wb = xw.Book(r'C:\Users\Benjamin\Documents\Projets_Python\cours_Or.xlsx')
sht = wb.sheets['Sheet1']
#sht.range('A1').value = 'Test'
derniereLigne = sht.range('A1').end('down').row


r = requests.get("https://www.joubert-change.fr/or-investissement/cours/piece-or-72-10-f-napoleon.html")
soup = BeautifulSoup(r.content,'lxml',from_encoding='utf-8')
itemNapoleon = soup.find_all('tr')
ligneNapoleon = itemNapoleon[2]
joubertAchete = ligneNapoleon.contents[5].contents[0]
joubertVend = ligneNapoleon.contents[13]
joubertVend = joubertVend.contents[0]
joubertAchete.replace('€','')
joubertVend.replace('€','')

#date
sht.range(derniereLigne+1,1).value = str(aujourdhui)


#Colonne joubert Achete
joubertAchete = str(joubertAchete).replace('€','')
joubertAchete = joubertAchete.replace(',00','')
sht.range(derniereLigne+1,2).value = int(joubertAchete)

#Colonne joubert Vend
joubertVend = str(joubertVend).replace('€','')
joubertVend = joubertVend.replace(',00','')

sht.range(derniereLigne+1,3).value = int(joubertVend)

print(derniereLigne)

wb.save()
wb.close()

os.system('wmic process where name="EXCEL.EXE" delete')
