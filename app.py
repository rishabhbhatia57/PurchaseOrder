from flask import Flask, jsonify, request
import json
from flask_cors import CORS
from main import scriptStarted, downloadFiles, scriptEnded, checkFolderStructure, mergeExcelsToOne
from pdfToTable import pdfToTable,getFilesToProcess
from colSplitting import colSplitting
from pivotTable import mergeToPivot
from generatingPackingSlip import generatingPackaingSlip
import os
from datetime import datetime
import log
import sys
import base64


ClientCode = {
   "PL":"Pantaloons",
   "SSL": "Shoppers Stop Limited",
   "LSL":"Lifestyle Limited"
}
posource = sys.argv[3].replace('#', ' ')
clientcode = sys.argv[1].replace('#', ' ')
orderdate =  sys.argv[2].replace('#', ' ')

print("\n"+orderdate+"\n"+"\n"+clientcode+"\n"+"\n"+posource+"\n")

clientname = ClientCode[clientcode]
destinationpath = "D:/ProjectStructure"
logger = log.setup_custom_logger('root')
formulasheetpath = '../PO Metadata/Formulasheet-Folder'
masterspath = '../PO Metadata/Masterfiles-Folder'
configpath = '../PO Metadata/Configfiles-Folder'
templatespath = '../PO Metadata/Templatesheet-Folder'


if __name__ == "__main__":
    '''
    01 DOWNLOAD FILES - FROM GGOGLE DRIVE - done
    02 EXTRACT CSV - INTERMIEIATE - PDF TO CSV - 
    03 EXTRACT EXCEL - CLEANING - CSV TO EXCEL
    04 CONSOLIDATE ORDER - MERGING PIVOT - PO SUMMERY.XSLS APPNDED - USE TEMPLATE TO MAKE 
    05 REQUIREMENT SUMMARY - PIVOT 
    06 PACKING SLIP(PIVOT TO SPLIT ORDERS -  GENRATE INDIVIDUAL PACKING SLIP ) - UISNG TEMPLATE PACKING SLIP - EXCLE PO TO ORDERDATA.ZLSX - IMP POSITION
    06 Upload packing slip into gooogle drive
    '''
    
# Phase I
# 1. Notify that the script is Started
    scriptStarted()
# 2. Checking the folder structure 
    checkFolderStructure(RootFolder=destinationpath,ClientCode=clientcode,OrderDate=orderdate)
# 3. To download PDF Files from Google Drive and Store it in week/DownloadFiles Folder
    downloadFiles(RootFolder=destinationpath,POSource=posource,OrderDate=orderdate,ClientCode=clientcode) # Done
# 4. Converted PDF files to Excel Files, perform Cleaning, and store to week/uploadFiles Folder
    getFilesToProcess(RootFolder=destinationpath,POSource=posource,OrderDate=orderdate,ClientCode=clientcode)
# 5. Merge all the coverted excel file to a single excel file and store in week/MergeExcelsFiles folder
    mergeExcelsToOne(RootFolder=destinationpath,POSource=posource,OrderDate=orderdate,ClientCode=clientcode)
# 6. PivotTable - Template Creation
    mergeToPivot(RootFolder=destinationpath,POSource=posource,OrderDate=orderdate,ClientCode=clientcode,Formulasheet=formulasheetpath)



# Phase II
    generatingPackaingSlip(RootFolder=destinationpath,POSource=posource,OrderDate=orderdate,ClientCode=clientcode,Formulasheet=formulasheetpath,TemplateFiles=templatespath)
# 7. Notify that the script is Ended
    scriptEnded()
