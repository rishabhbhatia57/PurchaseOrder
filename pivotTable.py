import pandas as pd
import numpy as np
import xlsxwriter
from datetime import datetime
from openpyxl import load_workbook,Workbook
import openpyxl.utils.cell
from datetime import datetime


def mergeToPivot():
    df = pd.read_excel("./Week/MergeExcelsFiles/Merged.xlsx")
    df_pivot = pd.pivot_table(df, index="ArticleEAN", values='Qty', 
    columns=['Vendor Name','PO Number','Receiving Location'], aggfunc='sum',margins=True,margins_name='Grand Total')
    df_pivot['Closing Stock'] = 0
    df_pivot['Diff CS - GT'] = 0
    df_pivot['Rate'] = 0
    df_pivot.to_excel('./Week/PivotTable/PivotTableoutput.xlsx')


    # Adding Processing Date, Order Number and Closing Stock, Diffrence Between Grand Total and 
    # Closing Stock Field into pivot sheet for tempalte
    pivotWorksheet = load_workbook('./Week/PivotTable/PivotTableoutput.xlsx')
    pivotSheet = pivotWorksheet.active

    df = pd.DataFrame(pivotSheet, index=None)
    rows = len(df.axes[0])
    cols = len(df.axes[1])

    

    for i in range(5,rows):
        pivotSheet.cell(i,cols-1).value = '='+openpyxl.utils.cell.get_column_letter(cols-2)+str(i+2)+'-'+openpyxl.utils.cell.get_column_letter(cols-3)+str(i+2)

    pivotSheet.insert_rows(1,2)
    pivotSheet.insert_rows(5) # IGST/CGST Type
    pivotSheet.cell(5,1).value = 'IGST/CGST Type'

    todayDate = datetime.today().strftime('%Y-%m-%d')
    pivotSheet.cell(2,3).value = 'Processing Date'
    pivotSheet.cell(2,4).value = todayDate
    pivotSheet.cell(2,6).value = 'Order Number'
    pivotSheet.cell(2,7).value = '-'


    pivotWorksheet.save("./Week/PivotTable/PivotTableoutput.xlsx")
    return 'Pivot Table generated.'

    # # Adding Processing Date, Order Number and Closing Stock, Diffrence Between Grand Total and 
    # # Closing Stock Field into pivot sheet for tempalte
    # df_pivot['Closing Stock'] = 0
    # # df_pivot['Diff GT - CS'] = df_pivot['Closing Stock'] - df_pivot['Grand Total']
    # df_pivot['Diff GT - CS'] = '=X5-W5'
    # df_pivot['Rate'] = 0
