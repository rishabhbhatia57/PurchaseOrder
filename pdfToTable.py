import tabula
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import csv
import os
import xlsxwriter

# inputPath = '4000441961.pdf'
# outputPath = "output.xlsx"


# os.remove(intermediateExcel)
# os.remove(intermediateCSV)
# os.remove(outputPath)

# Extracting data from pdf and storing in csv
def pdfToTable(inputPath, outputPath):
    lastrowofdata = 355
    intermediateCSV = 'csvoutput.csv'
    intermediateExcel = 'sheetoutput.xlsx'
    tabula.convert_into(input_path= inputPath , output_path= intermediateCSV , pages = 'all', lattice= True)

    # making a dataframe using csv
    df = pd.read_csv(filepath_or_buffer= intermediateCSV, skiprows = [1, 2, 3])


    # converting csv to excel
    writer = pd.ExcelWriter(intermediateExcel, engine='xlsxwriter')
    df.to_excel(writer, sheet_name= 'Sheet1', merge_cells= False)
    writer.save()


    # loading excel
    workbook = load_workbook(filename= intermediateExcel)
    sheet = workbook.active


    # finding the last row needed
    endrow = 1
    endcolumn = 1
    stop = False

    # sheet.cell(endrow, 1).value.find("AMOUNT")

    # while stop == False:
    #     if sheet.cell:
    #         endrow = endrow + 1
    #         print(endrow)
    #     stop = True
        

    # print(endpoint)
        


    # Cleaning loop
    for i in range(1, lastrowofdata):
        for j in range(1, 14):
            if sheet.cell(i,j).value != None and "Aditya" in sheet.cell(i,j).value:
                sheet.cell(i,1).value = "remove"
            elif sheet.cell(i,j).value != None and "P. O. Number" in sheet.cell(i,j).value:
                sheet.cell(i,1).value = "remove"
            elif sheet.cell(i,j).value != None and "PO" in sheet.cell(i,j).value and i != 2:
                    sheet.cell(i,1).value = "remove"
            elif sheet.cell(i,j).value != None and "_x000D_" in sheet.cell(i,j).value:
                temp = sheet.cell(i,j).value.replace("_x000D_", "")
                sheet.cell(i,j).value = temp




        

    # rearranging the data into a single row

    # column numbers
    CGSTRate = 14
    CGSTAmt = 15
    SGSTRate = 16
    SGSTAmt = 17
    UTGSTRate = 18
    UTGSTAmt = 19
    IGSTRate = 20
    IGSTAmt = 21
    VendorPenalty = 22
    for i in range(1, lastrowofdata):
        for j in range(1, 14):
            if sheet.cell(i,j).value == "CGST":
                sheet.cell(i-1,CGSTRate).value = sheet.cell(i,j+1).value
                sheet.cell(i-1,CGSTAmt).value = sheet.cell(i,j+2).value
                sheet.cell(i,j+1).value = ""
                sheet.cell(i,j+2).value = ""
                sheet.cell(i,j).value = ""
            elif sheet.cell(i,j).value == "SGST":
                sheet.cell(i-2,SGSTRate).value = sheet.cell(i,j+1).value
                sheet.cell(i-2,SGSTAmt).value = sheet.cell(i,j+2).value
                sheet.cell(i,j+1).value = ""
                sheet.cell(i,j+2).value = ""
                sheet.cell(i,j).value = ""
            elif sheet.cell(i,j).value == "UTGST":
                sheet.cell(i-3,UTGSTRate).value = sheet.cell(i,j+1).value
                sheet.cell(i-3,UTGSTAmt).value = sheet.cell(i,j+2).value
                sheet.cell(i,j+1).value = ""
                sheet.cell(i,j+2).value = ""
                sheet.cell(i,j).value = ""
            elif sheet.cell(i,j).value == "IGST":
                sheet.cell(i-4,IGSTRate).value = sheet.cell(i,j+1).value
                sheet.cell(i-4,IGSTAmt).value = sheet.cell(i,j+2).value
                sheet.cell(i,j+1).value = ""
                sheet.cell(i,j+2).value = ""
                sheet.cell(i,j).value = ""
            elif sheet.cell(i,j).value == "VendorPenalty":
                sheet.cell(i-4,VendorPenalty).value = sheet.cell(i,j+1).value
                sheet.cell(i,j+1).value = ""
                sheet.cell(i,j).value = ""


    # removing empty and unwanted rows
    counter = 354
    while counter > 0:
        if sheet.cell(counter,1).value == None or sheet.cell(counter,1).value == "" or sheet.cell(counter, 1).value == "remove":
            sheet.delete_rows(counter)
        counter = counter -1

    sheet["A1"] = "POItem"
    sheet["B1"] = "ArticleEAN"
    sheet["C1"] = "Article Number"
    sheet["D1"] = "ArticleDescription"
    sheet["E1"] = "HSNCode"
    sheet["F1"] = "MRP"
    sheet["G1"] = "BasicCostPrice(TaxableValue)"
    sheet["H1"] = "Qty"
    sheet["I1"] = "UM"
    sheet["J1"] = "TaxableValue"
    sheet["K1"] = "GSTRate"
    sheet["L1"] = "GSTAmt"
    sheet["M1"] = "Total Amount"
    sheet["N1"] = "CGSTRate"
    sheet["O1"] = "CGSTAmt"
    sheet["P1"] = "SGSTRate"
    sheet["Q1"] = "SGSTAmt"
    sheet["R1"] = "UTGSTRate"
    sheet["S1"] = "UTGSTAmt"
    sheet["T1"] = "IGSTRate"
    sheet["U1"] = "IGSTAmt"
    sheet["V1"] = "Vendor Penalty"



    workbook.save(filename= outputPath)
    print(outputPath)
    return "New File generated: "+str(outputPath)

