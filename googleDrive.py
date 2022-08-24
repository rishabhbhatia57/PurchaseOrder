import os
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from flask import Flask, redirect , url_for, render_template, request, jsonify
import json
gauth = GoogleAuth()
drive = GoogleDrive(gauth)

PO_Order_folder_ID = '1GcthE_OBicnjDxMdQQwIbXWVC2P750Pi'
PO_Packing_folder_ID = '189W-ssDT--oQvcT3tVmGCxD615b3_9Je'
PO_Summery_folder_ID = '1_vXO-Ks6EK26wMoUOQvSte7KhbaMnF7l'

def downloadFiles():
    # 1. Download files from google drive
        # get Files name: 
    file_list = drive.ListFile({'q': f"'{PO_Order_folder_ID}' in parents and trashed=false"}).GetList()
    # for f in file_list:
    #     print(f['title'])

        # Download files
    # List_of_Downloaded_files = []
    file_list = drive.ListFile({'q': f"'{PO_Order_folder_ID}' in parents and trashed=false"}).GetList()
    
    for index, f in enumerate(file_list):
        with open(os.path.join("./DownloadedFiles/",f['title']), 'wb') as c:
            f.GetContentFile("./DownloadedFiles/"+f['title'])  # Download files in downloadedfiles folder 
            c.close()
            print(f['title']+" downloaded from Google Drive!")
        # List_of_Downloaded_files.append(f['title'])
    
    return "Downloaded Complete!"
    # jsonify(
    #     Status = "Downloaded Complete!",
    #     List_of_Downloaded_files = List_of_Downloaded_files
    # )


def uploadFiles():
    # To upload files from a specific local directory
    #local directory path
    diretory = "./PO Order"
    #To list all the files in the local directory - to upload files from a specific directory
    for f in os.listdir(diretory):
        filename = os.path.join(diretory,f)
        gfile = drive.CreateFile({'parents' : [{'id' : PO_Packing_folder_ID}], 'title' : f})
        gfile.SetContentFile(filename)
        gfile.Upload()
        print(f+" uploaded to Google Drive!")

    # To upload a file
    # file1 = drive.CreateFile({'parents' : [{'id' : PO_Summery_folder_ID}], 'title' : 'hello.xlsx'})
    # file1.SetContentString('Hello world!') # set wrod/doc file content as the string
    # file1.Upload() # finally upload the file in google drive
    return "Files Uploaded successfully to google drive!"
    
    # jsonify(
    #     Status = "Files Uploaded successfully to google drive!"
    # )



# To upload files from a specific local directory
# #local directory path
# diretory = "C:/Users/HP/Desktop/TrimphPO/PO Order"
# To list all the files in the local directory - to upload files from a specific directory
# # for f in os.listdir(diretory):
#     filename = os.path.join(diretory,f)
#     print(filename)
#     gfile = drive.CreateFile({'parents' : [{'id' : PO_Packing_folder_ID}], 'title' : f})
#     gfile.SetContentFile(filename)
#     gfile.Upload()



# List files from google drive
# file_list = drive.ListFile({'q': f"'{PO_Order_folder_ID}' in parents and trashed=false"}).GetList()
# for f in file_list:
#     print(f['title'])


#Download files from google drive
# file_list = drive.ListFile({'q': f"'{PO_Order_folder_ID}' in parents and trashed=false"}).GetList()
# for index, f in enumerate(file_list):
#     # f.GetContentFile(f['title'])
#     print(index+1 ,"file downloaded: ", f['title'])
#     # print(f.auth.credentials.access_token)


# headers = {"access_token":"Bearer" + f.auth.credentials.access_token}

# def downloadFiles(File_ID):
#     downloadDirectoryPath = "C:/Users/HP/Desktop/TrimphPO/Downloaded Files"
#     # print("jhduabk\n")
#     response = request.get_data("https://www.googleapis.com/drive/v3/files/"+File_ID+"?alt=media&source=downloadUrl")
#     # .get("https://www.googleapis.com/drive/v3/files/"+File_ID+"?alt=media&source=downloadUrl")
#     print(response,"\n")
#     return jsonify(
#         Done = response.decode('utf-8')
#     )
    # print(response)
