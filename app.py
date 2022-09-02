from flask import Flask, jsonify, request
import json
from flask_cors import CORS
from googleDrive import downloadFiles, uploadFiles
from pdfToTable import pdfToTable,getFilesToProcess
from colSplitting import colSplitting
from mergeExcels import mergeExcelsToOne
from pivotTable import mergeToPivot
import os
from datetime import datetime


path = "./Week"

def scriptStarted():
    print("\nScript started...")
    print("Script is running...")
    return "Script Started."

def scriptEnded():
    print("Script Ended.\n")
    return "Script Ended"

def checkFolderStructure():
    todayDate = datetime.today().strftime('%Y-%m-%d')
    # Checking if the folder exists or not, if doesnt exists, then script will create one.
    DatedPath = path +"/"+todayDate
    isExist = os.path.exists(DatedPath)
    if not isExist:
        print("Creating a new folder '"+todayDate+"' at location "+ path)
        os.makedirs(DatedPath)
        os.makedirs(DatedPath+"/DownloadFiles")
        os.makedirs(DatedPath+"/MergeExcelsFiles")
        os.makedirs(DatedPath+"/PackagingSlip")
        os.makedirs(DatedPath+"/PivotTable")
        os.makedirs(DatedPath+"/ProcessedFiles")
        os.makedirs(DatedPath+"/IntermediateFiles")
        print("The new directory is created!")
    

    pass


if __name__ == "__main__":
    todayDate = datetime.today().strftime('%Y-%m-%d')
    '''
    01 DOWNLOAD FILES - FROM GGOGLE DRIVE - done
    02 EXTRACT CSV - INTERMIEIATE - PDF TO CSV - 
    03 eXTRACT eXCEL - CLEANING - CSV TO EXCEL
    04 CONSOLIDATE ORDER - MERGING PIVOT - PO SUMMERY.XSLS APPNDED - USE TEMPLATE TO MAKE 
    05 PACKING SLIP (PIVOT TO SPLIT ORDERS -  GENRATE INDIVIDUAL PACKING SLIP ) - UISNG TEMPLATE PACKING SLIP - EXCLE PO TO ORDERDATA.ZLSX - IMP POSITION
    06 Upload packing slip into gooogle drive
    '''
    
# 1. Notify that the script is Started
    scriptStarted()
    # checkFolderStructure()

# 2. To download PDF Files from Google Drive and Store it in week/DownloadFiles Folder
    # downloadFiles(path) # Done

# 3. Converted PDF files to Excel Files, perform Cleaning, and store to week/uploadFiles Folder
    # getFilesToProcess(path+"/DownloadFiles", path+"/ProcessedFiles")
    
# 4. Upload back converted Excel files to the google drive
    # uploadFiles(path+"DownloadFiles", path+"UploadFiles")
    
# 5. Split 2 columns and make seperate excel file - will be used for pivot tables
#    tale to 2 col from excel and make seprate file, this should be done for all the col in the  excel file
    # colSplitting(path+"PivotTable",path+"ColumnSplitting")
    
# 6. Merge all the coverted excel file to a single excel file and store in week/MergeExcelsFiles folder
    # mergeExcelsToOne(path+"/ProcessedFiles", path+"/MergeExcelsFiles")

# 7. PivotTable - Template Creation
    mergeToPivot()
    
# 7. Notify that the script is Ended
    scriptEnded()
