# -*- coding: utf-8 -*-
"""
Created on Tue Oct 17 01:10:22 2017

@author: Prosimio
"""
import numpy as np

# import data management package
import pickle as pkl

# Functions to load the data
def sheet_size(sheet, rc0=[1,1]): 
    """
    This function aims to size the data present in a sheet. The sizing is performed in a fixed position given by the user (rc0).
    It means if rc0=[1,1] the column counting will be done in the first row and the row counting will be done in the first column.
    
    Parameters
    ----------
    sheet: workbook sheet
        sheet file to be sized
    
    rc0: numeric list
        [row number, colmun number] fixed positions where perform the count
        
    
    Returns
    -------        
    rows: integer
        number of rows counted in the sheet
    
    columns: integer
        number of columns counted in the sheet
        
    """
    rows = 1 
    columns = 1
        
    while True:
        if sheet.cell(row=rows+1, column=rc0[1]).value != None: 
            rows+=1
        else:
            break
    while True:
        if sheet.cell(row=rc0[0], column=columns+1).value != None:
            columns+=1
        else:
            break
    #print(str(rows) + ' rows') 
    #print(str(columns) + ' columns')
    return(rows,columns)

def delimit_data():
    start_row=input('Enter the data starting row number (e.g. 2) ')
    print('data start in row '+ start_row) 
    start_column=input('Enter the data starting column number (e.g. 3)')
    print('data start in column '+ start_column)
    return(int(start_row), int(start_column))

def get_values(ws, row_start, column_start, time_column=1, header_row=1):
    """
    To get the wells data from a worksheet
    
    Parameters
    ----------
    ws: worksheet
        sheet file to get the data
    
    row_start: int
        row where wells data started
        
    
    Returns
    -------
    data: numpy array
        wells experimental measurements
        
    time: numpy vector
        time points values in hours units
    
    header: str
        header of each well data (usually 'A1','A2', etc)
        
    """
    
    rows, columns=sheet_size(ws)    
    data= np.zeros((rows-row_start+1,columns-column_start+1))
    time = np.zeros((rows-row_start+1))
    headers = []
    for i in range(rows-row_start+1):
        
        time_aux = ws.cell(row=i+row_start, column=time_column).value
        
        #transform time values in hour fractions units
        if type(time_aux)==float:
            time[i] = time_aux*24
        else:
            time[i] =time_aux.hour + time_aux.minute/60 + time_aux.second/3600
            
                
        for j in range(columns-column_start+1):
            data[i,j]=ws.cell(row=i+row_start, column=j+column_start).value
            if i==0:
                headers.append(ws.cell(row=header_row, column=j+column_start).value)  # list with the headers
    #print(headers)
    return(data,time,headers)

def delimit_label():
    start_row=input('Enter the label starting row number (e.g. 2) ')
    print('data start in row '+ start_row)
    start_column=input('Enter the label starting column number (e.g. 2) ')
    print('data start in column '+ start_column)
    
    return(int(start_row), int(start_column))

def get_label_values(ws,row_start,column_start):
    rows,columns=sheet_size(ws)    
    if type(ws.cell(row=row_start, column=column_start).value) == str:
        values = []
        for i in range(rows-row_start+1):
            values_col=[]         
            for j in range(columns-column_start+1):
                values_col.append(ws.cell(row=i+row_start, column=j+column_start).value)
            values.append(values_col)  
    
    else:   #in case data is not string
        values= np.zeros((rows-row_start+1,columns-column_start+1))
    
        for i in range(rows-row_start+1):
                      
            for j in range(columns-column_start+1):
                values[i,j]=ws.cell(row=i+row_start, column=j+column_start).value

    return(values)    

def save_obj(obj, name, folder ):
    with open(folder+'/'+ name + '.pkl', 'wb') as f:
        pkl.dump(obj, f, pkl.HIGHEST_PROTOCOL)

#####
