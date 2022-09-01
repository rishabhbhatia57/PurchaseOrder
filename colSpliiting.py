import pandas as pd
import os
import xlsxwriter

def colSplitting(inputpath, outputpath):
    isExist = os.path.exists(inputpath)
    if not isExist:
        print("Can't find the ProcessedFiles folder, please check the file structure again!")
        return
    else:
        isExist = os.path.exists(outputpath)
        if not isExist:
            print("Creating a new folder 'ColumnSplitting' at location "+ outputpath)
            os.makedirs(outputpath)
            print("The new directory is created!")
        for f in os.listdir(inputpath):
            print("Accessing '"+f+"' right now: ")
            df = pd.read_excel(inputpath+"/"+f)
            ListofCol = list(df.columns)
            for i in range(1,len(ListofCol)):   
                df2 = df.loc[:, ['ArticleEAN', ListofCol[i]]]
                df2.to_excel(outputpath+"/"+str(ListofCol[i])+" "+f, index = False)
            print("Column Splitting for file "+f+" is comepleted.")

    return "Splitting Completed"

