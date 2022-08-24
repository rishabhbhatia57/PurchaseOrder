import pandas as pd
import os
import xlsxwriter

def colSplitting():
    getFilesDir = "./PO Order/"
    for f in os.listdir("./PO Order"):
        print("\nAccessing '"+f+"' right now: \n")
        df = pd.read_excel(getFilesDir+f)
        ListofCol = list(df.columns)
        for i in range(1,len(ListofCol)):
            df2 = df.loc[:, ['POItem', ListofCol[i]]]
            df2.to_excel("./ColumnSplitting/POItem"+"_"+ListofCol[i]+" "+f, index = False)
            print("\n\t"+str(i)+" Created splitted colmuns: POItem_"+ListofCol[i]+"\n")

    return "Splitting Completed"