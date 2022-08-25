from flask import Flask, jsonify, request
import json
from flask_cors import CORS
from googleDrive import downloadFiles, uploadFiles
from pdfToTable import pdfToTable,getFilesToProcess
import os

# app = Flask(__name__)
# cors = CORS(app)
path = "./Week/"
# @app.route('/')
def scriptStarted():
    print("\nScript started...")
    print("Script is running...")
    return "Script Started"

def scriptEnded():
    print("Script Ended successfully!\n")
    return "Script Ended"


if __name__ == "__main__":
    scriptStarted()
    downloadFiles(path)
    getFilesToProcess(path+"DownloadFiles", path+"UploadFiles")
    uploadFiles(path+"UploadFiles")
    scriptEnded()
    # app.run(debug=True)