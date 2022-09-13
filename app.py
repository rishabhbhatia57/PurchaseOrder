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
import json
import base64



ClientCode = {
   "PL":"Pantaloons",
   "SSL": "Shoppers Stop Limited",
   "LSL":"Lifestyle Limited"
}


with open('C:/Users/HP/Desktop/PO Metadata/Configfiles-Folder/config.json', 'r') as jsonFile:
    config = json.load(jsonFile)
    formulasheetpath = config['formulaFolder']
    masterspath = config['masterFolder']
    # configpath = config['formulaFolder']
    templatespath = config['templateFolder']
    destinationpath = config['targetFolder']


logger = log.setup_custom_logger('root')



if __name__ == "__main__":

    mode = sys.argv[1].replace('#', ' ')
    
    # Phase I
    if mode == 'consolidation':
        clientname = sys.argv[2].replace('#', ' ')
        orderdate =  sys.argv[3].replace('#', ' ')
        posource = sys.argv[4].replace('#', ' ')
        # print(clientname)
        # key_list = list(ClientCode.keys())
        # val_list = list(ClientCode.values())
        # position = val_list.index(clientname)
        clientcode = clientname
        logger.info("Client Name: "+clientname+" Client Code: "+clientcode+" Order Date: "+orderdate+" PO Folder Path: '"+posource+"'" )
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
        scriptEnded()

    if mode == 'packaging':
        clientcode = sys.argv[2].replace('#', ' ')
        orderdate =  sys.argv[3].replace('#', ' ')
        reqsource = sys.argv[4].replace('#', ' ')
        clientname = ClientCode[clientcode]
        logger.info("Client Name: "+clientname+" Client Code: "+ClientCode+" Order Date: "+orderdate+" PO Folder Path: '"+reqsource+"'" )
    # Phase II
        scriptStarted()
        generatingPackaingSlip(RootFolder=destinationpath,ReqSource=reqsource,OrderDate=orderdate,ClientCode=clientcode,Formulasheet=formulasheetpath,TemplateFiles=templatespath)
    # 7. Notify that the script is Ended
        scriptEnded()
