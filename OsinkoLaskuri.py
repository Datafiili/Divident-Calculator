import openpyxl
from openpyxl import Workbook

wb = Workbook()

#Stocks that you don't want to include. For example non-divident stocks.
BlackList = ['OP-AASIA INDEKSI A','OP-AMERIKKA INDEKSI A','OP-EUROOPPA INDEKSI A']

Events = []
def ReadInvestments():
    # Define variable to load the dataframe
    dataframe = openpyxl.load_workbook("Sijoitukset.xlsx")
     
    # Define variable to read sheet
    worksheet1 = dataframe.active
     
    # Iterate the loop to read the cell values
    for row in range(1, worksheet1.max_row ): #worksheet1.max_row
        L = []
        for col in worksheet1.iter_cols(0, worksheet1.max_column):
            L.append(col[row].value)
        L.append(0)
        if L[0] not in BlackList:
            Events.append(L)

def printAll():
    for E in Events:
        #if(E[0] == "SSAB B"):
            print(E)        

def ClearBadData():
    for i in range(len(Events)):
        if Events[i][3] != None:
            Events[i][3] = Events[i][3].replace("kpl","")
            Events[i][3] = Events[i][3].replace(",",".")
            Events[i][3] = float("".join(Events[i][3].split()))

        if Events[i][5] != None:
            Events[i][5] = Events[i][5].replace("EUR","")
            Events[i][5] = Events[i][5].replace(",",".")
            Events[i][5] = float("".join(Events[i][5].split()))
        if Events[i][6] != None:
            Events[i][6] = Events[i][6].replace("EUR","")
            Events[i][6] = Events[i][6].replace(",",".")
            Events[i][6] = float("".join(Events[i][6].split()))

        #needs to be last, because non-euros require values from other cells
        if Events[i][4] != None:
            Events[i][4] = Events[i][4].replace(",",".")
            if "EUR" not in Events[i][4]: #Converts non-euros to euros
                Events[i][4] = float(Events[i][6]) / float(Events[i][3])
            else:
                Events[i][4] = Events[i][4].replace("EUR","")    
                Events[i][4] = float("".join(Events[i][4].split()))

def myFunc(e):
  return e[0]

def Sort():
    global Events
    Events = sorted(Events, key=lambda x: (x[0], x[2]))

def CalculateDividend():
    #Taxes
    for i in range(len(Events)):
        if Events[i][1] == "VERON PIDÄTYS":
            if Events[i-1][1] == "OSINKO" or Events[i][0] == Events[i-1][0]:
                PerOsake = float(Events[i][6]) / float(Events[i][3])
                Events[i-1][4] -= PerOsake
            else:
                Print("ERROR at Events[" + str(i) + "]")
    CurrentInvestment = ""
    for i in range(len(Events)):
        if Events[i][0] != CurrentInvestment:
            CurrentInvestment = Events[i][0]
        if Events[i][1] == "OSINKO":
    
            

def SetData():
    worksheet2 = wb.active
    worksheet2.title = "Results"
    Titles = ['Laji:','Tapahtuma:','Päivämäärä:','Määrä:','Kurssi:','Kulut:','Yhteensä:','Osinko:']
    for i in range(8):
        worksheet2.cell(row=1, column=i+1, value=Titles[i])
    skipped = 0
    for E in range(len(Events)):
        #if Events[E][1] == "OSTO":
            for i in range(8):
                worksheet2.cell(row=E+2-skipped, column=i+1, value=Events[E][i])
#         else:
#             skipped += 1
    wb.save('Sijoitukset2.xlsx')  

ReadInvestments()
ClearBadData()
Sort()
#printAll()
CalculateDividend()
SetData()
