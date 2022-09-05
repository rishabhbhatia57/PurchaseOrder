import pandas as pd
import os
import xlsxwriter
import time
from openpyxl import load_workbook,Workbook
import openpyxl.utils.cell
import shutil




def generatingPackingSlip():
    startedTemplating = time.time()
    sourcePivot = "./Week/PivotTable/PivotTableoutput.xlsx"
    source = "./Week/RequiredFiles/PackingSlip-Template.xlsx"
    destination = "./Week/RequiredFiles/TemplateFile.xlsx"
    

    # Making Copy of template file
    # shutil.copy(source, destination)
    # print("File copied successfully.")

    # Load work vook and sheets
    InputWorkbook = load_workbook(sourcePivot,data_only=True)
    # TemplateWorkbook = load_workbook(destination)

    InputSheet = InputWorkbook.active
    # TemplateSheet = TemplateWorkbook['ORDER']

    # Get rows and Column count
    df = pd.DataFrame(InputSheet, index=None)
    rows = len(df.axes[0])
    cols = len(df.axes[1])


    formulaWorksheet = load_workbook('./Week/Master Files/FormulaSheet.xlsx',data_only=True)
    formulaSheet = formulaWorksheet['FormulaSheet']
    
    
    
    for column in range(2,cols-3):
        startedTemplatingFile = time.time()
        # Making Copy of template file
        shutil.copy(source, destination)
        print("File copied successfully.")

        # Load work vook and sheets
        TemplateWorkbook = load_workbook(destination, data_only=True)
        TemplateSheet = TemplateWorkbook['ORDER']
        dbfsheet = TemplateWorkbook['DBF']

        
        TemplateSheet.cell(6,4).value = InputSheet.cell(7,2).value # Order Name
        TemplateSheet.cell(5,1).value = InputSheet.cell(7,2).value # Order Name
        # PO Number
        filename = InputSheet.cell(4,column).value
        TemplateSheet.cell(5,2).value = InputSheet.cell(4,column).value
        TemplateSheet.cell(6,2).value = InputSheet.cell(4,column).value
        TemplateSheet.cell(6,1).value = InputSheet.cell(4,column).value
        TemplateSheet.cell(6,3).value = InputSheet.cell(4,column).value
        # Receving Location
        TemplateSheet.cell(5,3).value = InputSheet.cell(5,column).value
        TemplateSheet.cell(4,2).value = InputSheet.cell(5,column).value

        TemplateSheet.cell(1,3).value = 'SGST/IGST'
        TemplateSheet.cell(1,4).value = InputSheet.cell(6,column).value # IGST/SGST Type
        print(TemplateSheet.cell(1,4).value,InputSheet.cell(6,column).value)

        # Copy EAN to template sheet
        Trows = 8
        Tcols = 5
        dbfrows = 2
        for row in range(7,rows):
            # if InputSheet.cell(row,column).value != None or InputSheet.cell(row,column).value != "":
            if str(InputSheet.cell(row,column).value).isnumeric():
                    
                # Copy Qty to template sheet
                TemplateSheet.cell(Trows,Tcols).value = InputSheet.cell(row,column).value 
                TemplateSheet.cell(Trows,Tcols+1).value = InputSheet.cell(row,column).value
                
                # Copy EAN to template sheet
                TemplateSheet.cell(Trows,Tcols-3).value = InputSheet.cell(row,1).value

                # VLOOKUP
                # StyleName
                TemplateSheet.cell(Trows,Tcols-4).value = "="+formulaSheet.cell(3,2).value.replace("#VAL#",str(Trows))

                # style
                TemplateSheet.cell(Trows,Tcols-2).value =  "="+formulaSheet.cell(4,2).value.replace("#VAL#",str(Trows))

                # SADM SKU
                TemplateSheet.cell(Trows,Tcols-1).value = "="+formulaSheet.cell(5,2).value.replace("#VAL#",str(Trows))
                
                # Rate (in Rs.) Order file
                TemplateSheet.cell(Trows,Tcols+3).value = "="+formulaSheet.cell(6,2).value.replace("#VAL#",str(Trows))

                # Closing stk
                TemplateSheet.cell(Trows,Tcols+5).value = "="+formulaSheet.cell(11,2).value.replace("#VAL#",str(Trows)) 

                #Cls stk vs order
                # TemplateSheet.cell(Trows,Tcols+6).value = TemplateSheet.cell(Trows,Tcols+5).value - TemplateSheet.cell(Trows,Tcols+2).value
                TemplateSheet.cell(Trows,Tcols+6).value = "="+openpyxl.utils.cell.get_column_letter(Tcols+5)+str(Trows) +'-'+openpyxl.utils.cell.get_column_letter(Tcols+2)+str(Trows) 

                # LOCATION2
                TemplateSheet.cell(Trows,Tcols+7).value = "="+formulaSheet.cell(7,2).value.replace("#VAL#",str(Trows))

                #BULK  / DTA  BULK  /  EOSS LOC
                TemplateSheet.cell(Trows,Tcols+8).value = "="+formulaSheet.cell(8,2).value.replace("#VAL#",str(Trows))

                #MRP
                TemplateSheet.cell(Trows,Tcols+9).value = "="+formulaSheet.cell(9,2).value.replace("#VAL#",str(Trows))

                # Adding values to DBF
                
                dbfsheet.cell(dbfrows, 1).value = '='+formulaSheet.cell(13,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Vouchertypename
                dbfsheet.cell(dbfrows, 2).value = '='+formulaSheet.cell(14,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CSNNO
                dbfsheet.cell(dbfrows, 3).value = "=ORDER!L4" #Date
                dbfsheet.cell(dbfrows, 4).value = "=ORDER!A5" #Reference
                dbfsheet.cell(dbfrows, 5).value = 0 #REF1  
                dbfsheet.cell(dbfrows, 6).value = "=ORDER!C5" #Dealer Name
                # dbfsheet.cell(dbfrows, 7).value = "=ORDER!C"+str(Trows) # PriceLevel/Style
                dbfsheet.cell(dbfrows, 7).value = '='+formulaSheet.cell(16,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) # PriceLevel
                dbfsheet.cell(dbfrows, 8).value = "=ORDER!D"+ str(Trows) #ItemName/ SKU
                dbfsheet.cell(dbfrows, 9).value = '='+formulaSheet.cell(17,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #GODOWN
                dbfsheet.cell(dbfrows, 10).value = "=SUMIF(ORDER!$D$8:$D$22298,DBF!$H"+str(dbfrows)+",ORDER!$G$8:$G$22298)" #Qty
                dbfsheet.cell(dbfrows, 11).value = "=VLOOKUP($H"+str(dbfrows)+",ORDER!D:H,5,FALSE)" #Rate
                dbfsheet.cell(dbfrows, 12).value = "=ROUND(J"+str(dbfrows)+"*K"+str(dbfrows)+",2)" #SUBTOTAL
                dbfsheet.cell(dbfrows, 13).value = '='+formulaSheet.cell(18,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #DISCPERC
                dbfsheet.cell(dbfrows, 14).value = "=ROUND(L"+str(dbfrows)+"*M"+str(dbfrows)+"/100,2)" #DISCAMT
                dbfsheet.cell(dbfrows, 15).value = "=L"+str(dbfrows)+"-N"+str(dbfrows) #ITEMVALUE
                dbfsheet.cell(dbfrows, 16).value = '=IF(R'+str(dbfrows)+'="Corsetry",IF(IFERROR((O'+str(dbfrows)+'/J'+str(dbfrows)+'),0)<1000,"CC Sales- Corsetry-Wholesale - IGST 5%","CC Sales- Corsetry-Wholesale - IGST 12%"))' #LedgerAcct
                dbfsheet.cell(dbfrows, 17).value = '='+formulaSheet.cell(19,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY1
                dbfsheet.cell(dbfrows, 18).value = '='+formulaSheet.cell(20,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT1
                dbfsheet.cell(dbfrows, 19).value = '='+formulaSheet.cell(21,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY2
                dbfsheet.cell(dbfrows, 20).value = '='+formulaSheet.cell(22,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT2
                dbfsheet.cell(dbfrows, 21).value = '='+formulaSheet.cell(23,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY3
                dbfsheet.cell(dbfrows, 22).value = '='+formulaSheet.cell(24,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT3
                dbfsheet.cell(dbfrows, 23).value = '='+formulaSheet.cell(25,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY4
                dbfsheet.cell(dbfrows, 24).value = '='+formulaSheet.cell(26,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT4
                dbfsheet.cell(dbfrows, 25).value = "=+O"+str(dbfrows) #ITEMTOTAL
                dbfsheet.cell(dbfrows, 26).value = "=+J"+str(dbfrows) #TOTALQTY
                dbfsheet.cell(dbfrows, 27).value = '='+formulaSheet.cell(27,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CDISCHEAD
                dbfsheet.cell(dbfrows, 28).value = '='+formulaSheet.cell(28,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CDISCPERC
                dbfsheet.cell(dbfrows, 29).value = "=ROUND(Y"+str(dbfrows)+"*AB"+str(dbfrows)+"/100,2)" #COMMONDISC
                dbfsheet.cell(dbfrows, 30).value = "=ROUND(SUM(Y"+str(dbfrows)+",0)-AC"+str(dbfrows)+",2)" #BEFORETAX
                dbfsheet.cell(dbfrows, 31).value = "" #TAXHEAD
                dbfsheet.cell(dbfrows, 32).value = "" #TAXPERC
                dbfsheet.cell(dbfrows, 33).value = "" #TAXAMT
                dbfsheet.cell(dbfrows, 34).value = "" #STAXHEAD
                dbfsheet.cell(dbfrows, 35).value = "" #STAXPERC
                dbfsheet.cell(dbfrows, 36).value = "" #STAXAMT
                dbfsheet.cell(dbfrows, 37).value = '=IF(RIGHT(P'+str(dbfrows)+',2)="5%","Output IGST - 5% - Tamilnadu","Output IGST - 12% - Tamilnadu")' #ITAXHEAD
                dbfsheet.cell(dbfrows, 38).value = '=IF(AK'+str(dbfrows)+'="Output IGST - 12% - TamilNadu",12,5)' #ITAXPERC
                dbfsheet.cell(dbfrows, 39).value = "=ROUND(SUM(AD"+str(dbfrows)+"*AL"+str(dbfrows)+"/100,0),2)" #ITAXAMT
                dbfsheet.cell(dbfrows, 40).value = "=+AP"+str(dbfrows) #NETAMT
                dbfsheet.cell(dbfrows, 41).value = 0.00 #ROUND
                dbfsheet.cell(dbfrows, 42).value = "=+(Y"+str(dbfrows)+"-AC"+str(dbfrows)+")+AM"+str(dbfrows) #ROUND1
                dbfsheet.cell(dbfrows, 43).value = '='+formulaSheet.cell(29,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #REFTYPE
                dbfsheet.cell(dbfrows, 44).value = "=B"+str(dbfrows) #Name
                dbfsheet.cell(dbfrows, 45).value = "=AN"+str(dbfrows) #REFAMT
                dbfsheet.cell(dbfrows, 46).value = '=CONCATENATE($B$2," ",$C$2," ","QTY",  " ", $Z$2)' #Narration
                dbfsheet.cell(dbfrows, 47).value = '='+formulaSheet.cell(30,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Transport
                dbfsheet.cell(dbfrows, 48).value = '='+formulaSheet.cell(31,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #transmode
                dbfsheet.cell(dbfrows, 49).value = '='+formulaSheet.cell(32,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #pymtterm
                dbfsheet.cell(dbfrows, 50).value = "=ORDER!B5" #OrderNo/PO number
                dbfsheet.cell(dbfrows, 51).value = "=ORDER!L4" #orddate
                dbfsheet.cell(dbfrows, 52).value = '='+formulaSheet.cell(33,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #DANO
                dbfsheet.cell(dbfrows, 53).value = '='+formulaSheet.cell(34,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd1
                dbfsheet.cell(dbfrows, 54).value = '='+formulaSheet.cell(35,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd2
                dbfsheet.cell(dbfrows, 55).value = '='+formulaSheet.cell(36,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd3
                dbfsheet.cell(dbfrows, 56).value = '='+formulaSheet.cell(37,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd4






                # print(dbfrows, Trows)
                
                Trows += 1
                dbfrows += 1 

        
        TemplateWorkbook.save("./Week/PackagingSlip/"+"PackagingSlip_"+str(filename)+".xlsx")

        print("Packing template Generated for: "+str(filename)+ " file - {:.2f}".format(time.time() - startedTemplatingFile,2)+ " seconds.")

    print("Total time taken {:.2f}".format(time.time() - startedTemplating,2)+ " seconds.")
    return 'Completed!'

generatingPackingSlip()