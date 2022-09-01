import pandas as pd
import os
import xlsxwriter
import glob


def mergeExcelsToOne(inputpath, outputpath):
    isExist = os.path.exists(inputpath)
    if not isExist:
        print("Can't find the ProcessedFiles folder, please check the file structure again!")
        return
    else:
        isExist = os.path.exists(outputpath)
        if not isExist:
            print("Creating a new folder 'MergeExcelsFiles' at location "+ outputpath)
            os.makedirs(outputpath)
            print("The new directory is created!")

        file_list =  glob.glob(inputpath+ "/*.xlsx")
        excl_list = []
        for f in os.listdir(inputpath):
            print("Accessing '"+f+"' right now: ")
            df = pd.read_excel(inputpath+"/"+f)
            df.insert(0, "file_name", f)
            # print(df)
            excl_list.append(df)
        excl_merged = pd.concat(excl_list, ignore_index=True,)
        excl_merged.to_excel(outputpath+"/"+'Merged.xlsx', index=False)
        print("Merging complete! ")
        return 'All excels are merged into a single excel file'

