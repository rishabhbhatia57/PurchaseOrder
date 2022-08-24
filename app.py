from flask import Flask, jsonify, request
import json
from flask_cors import CORS
from googleDrive import downloadFiles, uploadFiles
from pdfToTable import pdfToTable
import os

# app = Flask(__name__)
# cors = CORS(app)

# @app.route('/')
def scriptStarted():
    print("\nScript started...\n")
    print("\nScript is running...\n")
    return "Script Started"

def scriptEnded():
    print("Script Ended successfully...")
    return "Script Ended"

# @app.route('/download',)
def download():
    info = downloadFiles()
    print(info)
    return info

# @app.route('/upload',)
def uploaded():
    info = uploadFiles()
    return info

# @app.route('/Processing',)
def processing():
    directory = "./DownloadedFiles"
    outputPath = "./PO Order/"
    
    for f in os.listdir("./DownloadedFiles"):
        fres = f.replace('.pdf', '')
        info = pdfToTable(directory+"/"+f, outputPath=outputPath + fres + ".xlsx")
    return info

if __name__ == "__main__":
    scriptStarted()
    download()
    uploaded()
    scriptEnded()
    # app.run(debug=True)