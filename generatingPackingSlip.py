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
    DBFformula = formulaWorksheet['DBF']
    
    
    
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

        TemplateSheet.cell(1,1).value = 'DATE'
        TemplateSheet.cell(1,2).value = InputSheet.cell(2,4).value # Date

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
                
                dbfsheet.cell(dbfrows, 1).value = '='+DBFformula.cell(2,2).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Vouchertypename
                dbfsheet.cell(dbfrows, 2).value = '='+DBFformula.cell(2,3).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CSNNO
                dbfsheet.cell(dbfrows, 3).value = '='+DBFformula.cell(2,4).value.replace("#VAL#",str(Trows)) #Date
                dbfsheet.cell(dbfrows, 4).value = '='+DBFformula.cell(2,5).value.replace("#VAL#",str(Trows)) #Reference
                dbfsheet.cell(dbfrows, 5).value = '='+DBFformula.cell(2,6).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #REF1  
                dbfsheet.cell(dbfrows, 6).value = '='+DBFformula.cell(2,7).value.replace("#VAL#",str(Trows)) #Dealer Name
                dbfsheet.cell(dbfrows, 7).value = '='+DBFformula.cell(2,8).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) # PriceLevel
                dbfsheet.cell(dbfrows, 8).value = '='+DBFformula.cell(2,9).value.replace("#VAL#",str(Trows)) #ItemName/ SKU
                dbfsheet.cell(dbfrows, 9).value = '='+DBFformula.cell(2,10).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #GODOWN
                dbfsheet.cell(dbfrows, 10).value = '='+DBFformula.cell(2,11).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Qty
                dbfsheet.cell(dbfrows, 11).value = '='+DBFformula.cell(2,12).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Rate
                dbfsheet.cell(dbfrows, 12).value = '='+DBFformula.cell(2,13).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #SUBTOTAL
                dbfsheet.cell(dbfrows, 13).value = '='+DBFformula.cell(2,14).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #DISCPERC
                dbfsheet.cell(dbfrows, 14).value = '='+DBFformula.cell(2,15).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #DISCAMT
                dbfsheet.cell(dbfrows, 15).value = '='+DBFformula.cell(2,16).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #ITEMVALUE
                dbfsheet.cell(dbfrows, 16).value = '='+DBFformula.cell(2,17).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #LedgerAcct
                dbfsheet.cell(dbfrows, 17).value = '='+DBFformula.cell(2,18).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY1
                dbfsheet.cell(dbfrows, 18).value = '='+DBFformula.cell(2,19).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT1
                dbfsheet.cell(dbfrows, 19).value = '='+DBFformula.cell(2,20).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY2
                dbfsheet.cell(dbfrows, 20).value = '='+DBFformula.cell(2,21).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT2
                dbfsheet.cell(dbfrows, 21).value = '='+DBFformula.cell(2,22).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY3
                dbfsheet.cell(dbfrows, 22).value = '='+DBFformula.cell(2,23).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT3
                dbfsheet.cell(dbfrows, 23).value = '='+DBFformula.cell(2,24).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CATEGORY4
                dbfsheet.cell(dbfrows, 24).value = '='+DBFformula.cell(2,25).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COSTCENT4
                dbfsheet.cell(dbfrows, 25).value = '='+DBFformula.cell(2,26).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #ITEMTOTAL
                dbfsheet.cell(dbfrows, 26).value = '='+DBFformula.cell(2,27).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #TOTALQTY
                dbfsheet.cell(dbfrows, 27).value = '='+DBFformula.cell(2,28).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CDISCHEAD
                dbfsheet.cell(dbfrows, 28).value = '='+DBFformula.cell(2,29).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #CDISCPERC
                dbfsheet.cell(dbfrows, 29).value = '='+DBFformula.cell(2,30).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #COMMONDISC
                dbfsheet.cell(dbfrows, 30).value = '='+DBFformula.cell(2,31).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #BEFORETAX
                dbfsheet.cell(dbfrows, 31).value = "" #TAXHEAD
                dbfsheet.cell(dbfrows, 32).value = "" #TAXPERC
                dbfsheet.cell(dbfrows, 33).value = "" #TAXAMT
                dbfsheet.cell(dbfrows, 34).value = "" #STAXHEAD
                dbfsheet.cell(dbfrows, 35).value = "" #STAXPERC
                dbfsheet.cell(dbfrows, 36).value = "" #STAXAMT
                dbfsheet.cell(dbfrows, 37).value = '='+DBFformula.cell(2,38).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #ITAXHEAD
                dbfsheet.cell(dbfrows, 38).value = '='+DBFformula.cell(2,39).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #ITAXPERC
                dbfsheet.cell(dbfrows, 39).value = '='+DBFformula.cell(2,40).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #ITAXAMT
                dbfsheet.cell(dbfrows, 40).value = '='+DBFformula.cell(2,41).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #NETAMT
                dbfsheet.cell(dbfrows, 41).value = '='+DBFformula.cell(2,42).value.replace("#VAL#",str(Trows)) #ROUND
                dbfsheet.cell(dbfrows, 42).value = '='+DBFformula.cell(2,43).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #ROUND1
                dbfsheet.cell(dbfrows, 43).value = '='+DBFformula.cell(2,44).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #REFTYPE
                dbfsheet.cell(dbfrows, 44).value = '='+DBFformula.cell(2,45).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Name
                dbfsheet.cell(dbfrows, 45).value = '='+DBFformula.cell(2,46).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #REFAMT
                dbfsheet.cell(dbfrows, 46).value = '='+DBFformula.cell(2,47).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Narration
                dbfsheet.cell(dbfrows, 47).value = '='+DBFformula.cell(2,48).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Transport
                dbfsheet.cell(dbfrows, 48).value = '='+DBFformula.cell(2,49).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #transmode
                dbfsheet.cell(dbfrows, 49).value = '='+DBFformula.cell(2,50).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #pymtterm
                dbfsheet.cell(dbfrows, 50).value = '='+DBFformula.cell(2,51).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #OrderNo/PO number
                dbfsheet.cell(dbfrows, 51).value = '='+DBFformula.cell(2,52).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #orddate
                dbfsheet.cell(dbfrows, 52).value = '='+DBFformula.cell(2,53).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #DANO
                dbfsheet.cell(dbfrows, 53).value = '='+DBFformula.cell(2,54).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd1
                dbfsheet.cell(dbfrows, 54).value = '='+DBFformula.cell(2,55).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd2
                dbfsheet.cell(dbfrows, 55).value = '='+DBFformula.cell(2,56).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd3
                dbfsheet.cell(dbfrows, 56).value = '='+DBFformula.cell(2,57).value.replace("#VAL#",str(Trows)).replace("#DBFROWS#",str(dbfrows)) #Delyadd4

                




                # print(dbfrows, Trows)
                
                Trows += 1
                dbfrows += 1 

        
        TemplateWorkbook.save("./Week/PackagingSlip/"+"PackagingSlip_"+str(filename)+".xlsx")

        print("Packing template Generated for: "+str(filename)+ " file - {:.2f}".format(time.time() - startedTemplatingFile,2)+ " seconds.")

    print("Total time taken {:.2f}".format(time.time() - startedTemplating,2)+ " seconds.")
    return 'Completed!'

generatingPackingSlip()