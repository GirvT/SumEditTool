# -*- coding: utf-8 -*-
"""
Created on Mon Mar 12 10:01:01 2019

@author: Girvan Tse

"""
import re
from PySimpleGUI import Text, FileBrowse, Input, Window, Popup, Submit, Cancel, Checkbox, Button, Column
import numpy as np
from tkinter import TclError
from pandas import ExcelWriter, DataFrame, read_excel
from xlrd import XLRDError

layout =   [[Text('File to Query')],

            [Input('[Path to Excel Workbook]',
                  key = 'path'),
             FileBrowse(file_types=(("Excel Workbook",
                                     "*.xlsx"),
                                    ("All Files",
                                     "*.*")
                                    )),],
            [Input('[Sheet Name]',
                   key = 'sheet',
                   size = (54, 0))], 
            [Input('[Column Label]',
                  key = 'colName',
                  size = (32, 0)),
             Input('[Column Sum]',
                  key = 'colSum',
                  size = (20, 0))],
             [Input('[# of Sum Elements]',
                  key = 'numsums',
                  size = (20, 0))],
            [Submit(key = 'next'), Cancel(key = 'exit')]]

layout2 =  [[Text('Filter which elements?')]]

window = Window('RowRemoveTool' ).Layout(layout)

def validate(file):
    try:
        _testParam = read_excel(file[0], 
                                sheet_name = file[1])
    except FileNotFoundError:
        return 0
    except XLRDError:
        return 0
    return 1

RunTool = True
while RunTool:
    # Ignore TclError
    try:
        event, values = window.Read()
    except TclError:
        pass
    
    # Scan for specific input
    if (event is None or event == 'exit'):
        RunTool = False
    
    # Selecting row drop criteria
    if (event is 'next' and validate([values['path'], values['sheet']])):
        window.Close()
        colname = values['colName']
        colsum = values['colSum']
        try:
            numsums = int(values['numsums'])
        except:
            numsums = 0
        queryFrame = read_excel(values['path'],
                                sheet_name = values['sheet'])
        #drop useless columns
        dropCols = list()
        for column in queryFrame.columns:
            if (column.startswith('Unnamed: ')):
                dropCols.append(column)
        for column in dropCols:
            queryFrame.drop(column, axis=1, inplace=True)

        #Select how many suspect rows to drop
        PATH = values['path']
        rowList = list()
        #fix this
        for header in range(0, numsums):
            rowList.append([Input('[Name]')])
        layout2.append([Column(rowList, size = (400, 500), scrollable = True)])
        layout2.append([Submit(key = 'next1'), Cancel(key = 'exit')])
        window = Window('RowRemoveTool').Layout(layout2)

    # Drop suspect rows and export
    if (event is 'next1'):
        RunTool = False
        window.Close()


        SUM_FROM_COL = 7

        clec = queryFrame.select_dtypes(np.number).columns
        queryFrame[clec] = queryFrame[clec].fillna(0)
        queryFrame = queryFrame.fillna("")
        #drop all rows that meet criteria
        sumList = list(i for i in queryFrame.itertuples() if str(getattr(i, colname)) in list(values.values()))
        for tupl in sumList:
            rowTupl = tupl[0]
            sumTupl = 0
            for i in range(SUM_FROM_COL, len(tupl)):
                try:
                    sumTupl += tupl[i]
                except:
                    sumTupl += 0
            print(sumTupl)
            queryFrame.loc[rowTupl, colsum] = sumTupl
        #write results to output
        writer = ExcelWriter(PATH[:-5] + " OUTPUT.xlsx",
                             engine = 'xlsxwriter')
        queryFrame.to_excel(writer, sheet_name = 'Output', index = False)        
        writer.save()
        writer.close()

        #Let user know we're finished
        Popup('Successful Execution!')
