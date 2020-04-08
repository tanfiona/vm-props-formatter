import os
import pandas as pd

def read_excel(workbook_name, sheet_names=None):
    """
    Open an Excel spreadsheet and return the data in pandas.DataFrame format
    
    Parameters
    ----------
    workbook_name : str
        Excel spreadsheet filename
    sheet_names : list of str
        List of sheet names
    
    Returns
    -------
    data : dict of pandas.DataFrame
        Dictionary of data read from all sheets in the Excel spreadsheet
    """
    print('Reading Excel spreadsheet:', workbook_name)
    data = {}
    workbook = pd.ExcelFile(workbook_name)
    names = sheet_names if sheet_names is not None else workbook.sheet_names
    for name in names:
        print('Sheet:', name)
        data[name] = pd.read_excel(workbook, name)
    return data

def write_excel(dataframes, workbook_name, sheet_names=None, overwrite=False):
    """
    Write the pandas.DataFrame into an Excel spreadsheet
    
    Parameters
    ----------
    dataframes : dict of pandas.DataFrames
        Dictionary of data to be written to the Excel spreadsheet
    workbook_name : str
        Excel spreadsheet filename
    sheet_names : list of str
        List of sheet names
    overwrite : bool
        True to overwrite any existing spreadsheet or False if otherwise

    Returns
    -------
    None
    """
    print('Writing Excel spreadsheet:', workbook_name)
    writer = pd.ExcelWriter(workbook_name, engine='openpyxl')
    names = sheet_names if sheet_names is not None and len(sheet_names) == len(dataframes) else list(dataframes.keys())
    for name in names:
        print('Sheet:', name)
        dataframes[name].to_excel(writer, sheet_name=name, index=False)
    writer.save()
