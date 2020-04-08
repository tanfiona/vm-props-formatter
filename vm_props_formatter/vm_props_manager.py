import copy
import openpyxl
import pandas as pd
import numpy as np
#import glob
#import os
#from .utils.file_organizer import check_create_directory


class VMPropsManager(object):
    """"""
    __data_parser = None
    __parameters = None
    __defaults = {
        "shape": {
            "title_row": 4,
            "title_col": 1,
            "main_header_row": 10,
            "props_header_start_col": 5,
            "props_header_tally_first": 5,
            "no_summary_table_rows": 3,
            "summary_table_sum_row": 1,
            "props_header_end_col": -1,
        },
        "names": {
            "country_col": "COUNTRY NAME",
            "store_col": "STORE NAME",
            "main_cols": ["COUNTRY NAME", "STORE NAME"],
            "entity_list": ["CKS", "CKI", "CKC"]
        }
    }

    def __init__(self, parameters=None):
        """
        Constructor that copies the parameters

        Parameters
        ----------
        parameters : dict
            Dictionary of parameters

        Returns
        -------
        None
        """
        self.__parameters = copy.deepcopy(self.__defaults)
        # Set defined parameter values if present
        if parameters is not None:
            self.__parameters.update(parameters)

    def update_parameters(self, parameters):
        """
        Update the parameters

        Parameters
        ----------
        parameters : dict
            Dictionary of parameters

        Returns
        -------
        None
        """
        self.__parameters.update(parameters)

    def get_default_parameters(self):
        """
        Get the default parameters

        Parameters
        ----------

        Returns
        -------
        parameters : dict
            Default parameters
        """
        return self.__defaults

    def get_entity(self, file_name):
        for i in self.__parameters['names']['entity_list']:
            if i in file_name:
                return i

    def load_dataset(self, file_path, file_name, sheet_name = None):
        """
        Loads excel files into data and colour information
        
        Parameters
        ----------
        file_path: byte
        file_name: str
        sheet_name: str

        Returns
        -------
        data: pandas DataFrame
        sh : object

        """
        try:
            if sheet_name is None:
                sheet_name = self.get_entity(file_name)
                print('[Status] Sheet name not specified. Took based on file name entity name: "%s"' % sheet_name)
            #     sheet_name = pd.ExcelFile(file_path).sheet_names[0]
            #     print('[Status] Sheet name not specified. Took first sheet: ', sheet_name)
            print('[Status] Loading File: "%s";' % file_name, '"Loading Sheet: "%s"' % sheet_name)
            data = pd.read_excel(file_path, sheet_name, index_col=None, header=None)
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sh = wb[sheet_name]
            return data, sh
        except:
            print('[Error]  Incorrect file format, table columns, or table contents')
            return None, None

    def get_main_data(self, data, skiprows = None):
        """
        Shortens and re-header data

        Parameters
        ----------
        data: pandas.DataFrame
            Dataframe to run the operation on
        skiprows: int
            Number of rows to skip to first header row
            E.g. skiprows = 7 if the main data header is on Row '8'

        Returns
        -------

        """
        # load parameters if not specified
        if skiprows is None:
            skiprows = self.__parameters['shape']['main_header_row']
        # run analysis
        new_header = data.iloc[skiprows]
        data = data.iloc[skiprows+1:,]
        data.columns = new_header
        return data

    def dropna_rows_cols(self, data):
        """

        Parameters
        ----------
        data

        Returns
        -------

        """
        # drop rows that have na
        data = data.dropna(how='all')
        # get idx of last col with value
        last_idx = max(np.where(~(data.columns).isna())[0]) + 1
        # drop cols that have na AFTER main frame
        data = data.iloc[:, :last_idx]
        return data
