import pandas as pd
import os
import xlsxwriter
import glob


def mergeExcelsToOne():
    file_list =  glob.glob("./PO Order" + "/*.xlsx")
    excl_list = []
    for f in os.listdir("./PO Order"):
        print("\nAccessing '"+f+"' right now: \n")
        print("\nMerging '"+f+"' ... \n")
        df = pd.read_excel("./PO Order/"+f)
        df.insert(0, "file_name", f)
        # print(df)
        excl_list.append(df)
    excl_merged = pd.concat(excl_list, ignore_index=True,)
    excl_merged.to_excel("./MajorMergeExcel/"+'Mereged.xlsx', index=False)
    print("\nMerging complete! \n")
    return 'All excels are merged into a single excel file'

mergeExcelsToOne()