import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from flask import Flask, redirect , url_for, render_template, request, jsonify
import json
import time
gauth = GoogleAuth()
drive = GoogleDrive(gauth)

PO_Order_folder_ID = '1GcthE_OBicnjDxMdQQwIbXWVC2P750Pi'
PO_Packing_folder_ID = '189W-ssDT--oQvcT3tVmGCxD615b3_9Je'
PO_Summery_folder_ID = '1_vXO-Ks6EK26wMoUOQvSte7KhbaMnF7l'

def downloadFiles(path):
    startedDownload = time.time()
    # Checking if the folder exists or not, if doesnt exists, then script will create one.
    downloadPath = path + 'DownloadFiles'
    isExist = os.path.exists(downloadPath)
    if not isExist:
        print("Creating a new folder 'DownloadFiles' at location "+ path)
        os.makedirs(downloadPath)
        print("The new directory is created!")

    # Getting list of files pretent in google drive folder -  for downloading:
    file_list = drive.ListFile({'q': f"'{PO_Order_folder_ID}' in parents and trashed=false"}).GetList()    
    for index, f in enumerate(file_list):
        with open(os.path.join(downloadPath+"/",f['title']), 'wb') as c:
            f.GetContentFile(downloadPath+"/"+f['title'])  # Download files in DownloadFiles folder 
            c.close()
            print(""+f['title']+" downloaded from Google Drive!")
    print("Downloaded Completed in "+"{:.2f}".format(time.time() - startedDownload,2)+ " seconds!")
    
    return "Downloaded Complete!"


def uploadFiles(path):
    startedUpload = time.time()
    isExist = os.path.exists(path)
    if isExist == False:
        print("Can't find the UploadFiles folder, please check the file structure again!")
        return
    else:
        if len(os.listdir(path)) == 0:
            print("'"+path+"' Folder is empty, add files to upload.")
    #To list all the files in the local directory - to upload files from a specific directory
    for f in os.listdir(path):
        filename = os.path.join(path,f)
        gfile = drive.CreateFile({'parents' : [{'id' : PO_Packing_folder_ID}], 'title' : f})
        gfile.SetContentFile(filename)
        gfile.Upload()
        print(f+" uploaded to Google Drive!")
    print("Uploaded Completed in "+"{:.2f}".format(time.time() - startedUpload,2)+ " seconds!")
    return "Files Uploaded successfully to google drive!"
