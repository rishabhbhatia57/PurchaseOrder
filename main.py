import log
import os
import shutil
from datetime import datetime
import pandas as pd
import xlsxwriter
import glob


logger = log.setup_custom_logger('root')


def downloadFiles(RootFolder,POSource,OrderDate,ClientCode):
    #converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    source_folder = POSource
    destination_folder = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/10-Download-Files/"
    try:
        for file_name in os.listdir(POSource):
            # construct full file path
            
            source = source_folder +"/"+ file_name
            destination = destination_folder
            # copy only files
            if os.path.isfile(source):

                shutil.copy(source, destination)
                logger.info("File '"+file_name+"' copied from source '"+source_folder+"' to destination '"+destination_folder+"'")
    except Exception as e:
        logger.error("Error while copying files: "+str(e))

def scriptStarted():
    logger.info('Starting script')
    print('Starting script')
    print('Starting is running...')
    print('Do not close this window while processing...')
    return "Script Started."

def scriptEnded():
    logger.info('Script Ended')
    print('Script Ended')
    return "Script Ended"
    
def checkFolderStructure(RootFolder,ClientCode,OrderDate):
    try:
        #converting str to datetime
        OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')

        # extracting year from the order date
        year = OrderDate.strftime("%Y")

        # formatting order date {2022-00-00) format
        OrderDate = OrderDate.strftime('%Y-%m-%d')
        
        # Checking if the folder exists or not, if doesnt exists, then script will create one.
        logger.info('Checking if the folder exists or not, if doesnt exists, then script will create one.')
        DatedPath = RootFolder +"/"+ClientCode+"-"+year+"/"+str(OrderDate)
        isExist = os.path.exists(DatedPath)
        if not isExist:
            logger.info("Creating a new folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' at location "+ RootFolder)
            os.makedirs(DatedPath)
            os.makedirs(DatedPath+"/10-Download-Files")
            os.makedirs(DatedPath+"/20-Intermediate-Files")
            os.makedirs(DatedPath+"/30-Extract-CSV")
            os.makedirs(DatedPath+"/40-Extract-Excel")
            os.makedirs(DatedPath+"/50-Consolidate-Orders")
            os.makedirs(DatedPath+"/60-Requirement-Summary")
            os.makedirs(DatedPath+"/70-Packaging-Slip")
            os.makedirs(DatedPath+"/80-Logs")
            logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' is created.")
        logger.info("Folder '"+ClientCode+"-"+year+"/"+str(OrderDate)+"' exists.")

    except Exception as e:
        logger.error("Error while checking folder structure:  "+str(e))

def mergeExcelsToOne(RootFolder,POSource,OrderDate,ClientCode): 
    #converting str to datetime
    OrderDate = datetime.strptime(OrderDate, '%Y-%m-%d')
    # extracting year from the order date
    year = OrderDate.strftime("%Y")
    # formatting order date {2022-00-00) format
    OrderDate = OrderDate.strftime('%Y-%m-%d')
    inputpath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/40-Extract-Excel/"
    outputpath = RootFolder+"/"+ClientCode+"-"+year+"/"+OrderDate+"/50-Consolidate-Orders/"

    try:
        # logger.info('Checking 40-Extract-Excel directory exists or not.')
        file_list =  glob.glob(inputpath+ "/*.xlsx")
        if len(file_list) == 0:
            logger.info('No excel files found to merge.')
            return
        else:
            excl_list = []
            for f in os.listdir(inputpath):
                logger.info("Accessing '"+f+"' right now: ")
                df = pd.read_excel(inputpath+"/"+f)
                df.insert(0, "file_name", f)
                excl_list.append(df)
            excl_merged = pd.concat(excl_list, ignore_index=True,)
            excl_merged.to_excel(outputpath+"/"+'Consolidate-Orders.xlsx', index=False)
            logger.info("Merged "+str(len(file_list))+ " excel files to a single excel as 'Consolidate-Orders.xlsx'")
            return 'All excels are merged into a single excel file'
    except Exception as e:
        logger.info("Error while merging files: "+str(e))
