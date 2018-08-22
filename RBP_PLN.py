import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.cell import Cell

raport = Workbook()
#ws = raport.create_sheet()
ws = raport.active
ws.title = 'arkusz 1'

filename = r"C:/Users/Monika/Desktop/eoffice/RBP_PLN.xml"

tree = ET.parse(filename)
root = tree.getroot()

for child in root:
    #for header in child:
        #print(header.tag[48:len(header.tag)], len(header.tag))         #drukuje GrpHdr 54 i PmtInf 54    
    count_row = 1
    for line in child[1]:                                               #czyli tylko dla wiersza w PmtInf        
        #print(line.tag[48:len(line.tag)], len(line.tag))
        if line.tag[48:len(line.tag)] == "CdtTrfTxInf":
            #print(line.tag[48:len(line.tag)], len(line.tag))
            count_col = 1
            lista = []
            for cdttrf in line:
                p#rint(cdttrf.tag[48:len(cdttrf.tag)])
                try:
                    lista.append(cdttrf[0].text)
                    if count_col == 2:
                        kwota = cdttrf[0].text.replace('.',',')
                        ws.cell(column=count_col, row=count_row).value = kwota
                    else:
                        ws.cell(column=count_col, row=count_row).value = cdttrf[0].text
                except:
                    lista.append(0)
                    ws.cell(column=count_col, row=count_row).value = 0
                if cdttrf.tag[48:len(cdttrf.tag)] == 'CdtrAcct':    
                    ws.cell(column=count_col, row=count_row).value = cdttrf[0][0][0].text
                count_col +=1 
            count_row += 1
            #print('licznik',count_row)
            #print(lista)
                    
raport.save('RBP_PLN.xlsx')
                        
