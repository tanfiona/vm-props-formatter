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
            "main_header_row": 7,
            "props_header_start_col": 5,
            "props_header_tally_first": 5,
            "no_summary_table_rows": 3,
            "summary_table_sum_row": 1,
            "props_header_end_col": -1,
        },
        "names": {
            "sheet_name": "",
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

    def load_dataset(self, file_path, file_name, sheet_name=None):
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
        # load parameters if not specified
        if sheet_name is None:
            sheet_name = self.__parameters['names']['sheet_name']
        try:
            if sheet_name is "": # not specified by user either
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
        last_idx = max(np.where(~data.columns.isna())[0]) + 1
        # drop cols that have na AFTER main frame
        data = data.iloc[:, :last_idx]
        return data

    def get_index_to_split_tables2(self, data_1_clean, country_col=None):
        # load parameters if not specified
        if country_col is None:
            country_col = self.__parameters['names']['country_col']
        return [data_1_clean[~pd.isna(data_1_clean[country_col])].index[0]]

    def get_index_to_split_tables(self, main_data, skipcols=None, tallycols=None):
        """ 
        Function to find number of split tables based on first row header name
        
        Parameters
        ----------
        skipcols
        tallycols

        Returns
        -------

        """
        # load parameters if not specified
        if skipcols is None:
            skipcols = self.__parameters['shape']['props_header_start_col']
        if tallycols is None:
            tallycols = self.__parameters['shape']['props_header_tally_first']
        # run analysis
        col_name_list = []
        # get row indexes which tally with colnames
        sub_data = main_data.iloc[:, skipcols:skipcols + tallycols]
        for i in list(sub_data.index):
            if list(sub_data.loc[i].values) == list(sub_data.loc[i].index):
                col_name_list.append(i)
        return col_name_list

    # split data by col_name_list
    def get_split_data(self, data, col_name_list):
        # run analysis
        if len(col_name_list) == 1:
            data_1 = data.loc[:col_name_list[0]-1,]
            data_2 = data.loc[col_name_list[0]-1:,]
            return data_1, data_2
        elif len(col_name_list) == 2:
            data_1 = data.loc[:col_name_list[0]-1,]
            data_2 = data.loc[col_name_list[0]-1:col_name_list[1]-1,]
            data_3 = data.loc[col_name_list[1]-1:,]
            return data_1, data_2, data_3

    def clean_main_data(self, data, country_col=None):
        """
        
        Parameters
        ----------
        data
        country_col

        Returns
        -------

        """
        # load parameters if not specified
        if country_col is None:
            country_col = self.__parameters['names']['country_col']
        # run analysis
        # drop 'duplicate' rows where totals are dupl
        # hard-coded!!!!
        data = data[data['NO'] != 'Total:']
        # format country
        country_col_2_ind = data.columns.get_loc(country_col) + 1
        data[country_col] = data.iloc[:, country_col_2_ind].fillna(data[country_col])
        # fill na with above country
        for i in range(1, len(data)):
            value = data[country_col].iloc[i]
            prev_value = data[country_col].iloc[i - 1]
            if pd.isna(value):
                data[country_col].iloc[i] = prev_value
        return data

    def shorten_table_w_max_rows(self, data, max_rows=None):
        """
        
        Parameters
        ----------
        data: pandas.DataFrame
        max_rows: int

        Returns
        -------

        """
        # load parameters if not specified
        if max_rows is None:
            max_rows = self.__parameters['shape']['no_summary_table_rows']
        # run analysis
        return data.iloc[:max_rows,]

    def format_main_data(self, data):
        """
        
        Parameters
        ----------
        data

        Returns
        -------

        """
        # fillna
        df = data.fillna(0)

        # clear white space cells
        df = df.applymap(lambda x: 0 if str(x).isspace() else x)

        # add row sum
        df.loc[max(df.index) + 1] = df.sum(numeric_only=True)

        # fillna
        df = df.fillna(0)

        return df

    def main_and_summary_checker(self, df, summary, sum_row=None, skipcols_front=None, skipcols_end=None):
        # load parameters if not specified
        if sum_row is None:
            sum_row = self.__parameters['shape']['summary_table_sum_row']
        if skipcols_front is None:
            skipcols_front = self.__parameters['shape']['props_header_start_col']
        if skipcols_end is None:
            skipcols_end = self.__parameters['shape']['props_header_end_col']
        # run analysis
        checker = pd.DataFrame()
        checker['main'] = list(df.iloc[-1, skipcols_front:skipcols_end].fillna(0).astype(int))
        checker['summary'] = list(summary.iloc[sum_row, skipcols_front:skipcols_end].fillna(0).astype(int))
        checker['checks'] = list(checker['main'] == checker['summary'])
        checker = checker.set_index(df.iloc[:, skipcols_front:skipcols_end].columns)
        return checker

    def get_excel_col_from_int(self, n):
        string = ""
        n = int(n)
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string
    
    def main_table_to_so_converter(self, df, country_col=None, skipcols_front=None, skipcols_end=None):
        # load parameters if not specified
        if country_col is None:
            country_col = self.__parameters['names']['country_col']
        if skipcols_front is None:
            skipcols_front = self.__parameters['shape']['props_header_start_col']
        if skipcols_end is None:
            skipcols_end = self.__parameters['shape']['props_header_end_col']
        # run analysis
        # get prop column name list from column location
        props_column_names = list(df.iloc[:, skipcols_front:skipcols_end].columns)
        # melt all props into country
        df['XRow'] = df.index + 1
        so_table = pd.melt(df[['XRow', country_col] + props_column_names], id_vars=['XRow', country_col],
                           value_vars=props_column_names)
        # rename column names to fixed format
        so_table.columns = ['XRow', country_col, 'VM PROPS', 'Qty']
        # order columns
        so_table = so_table[[country_col, 'VM PROPS', 'Qty', 'XRow']]
        # drop props with no qty
        so_table['Qty'] = pd.to_numeric(so_table['Qty'], errors='coerce')
        so_table = so_table[so_table['Qty'] > 0]
        # rename total rows
        so_table.loc[so_table[country_col] == 0, country_col] = 'TOTAL'
        so_table.loc[so_table[country_col] == 0, country_col] = '0'
        # calculate cell location (of original excel)
        max_rows = len(df)
        max_ind = max(df.index)
        min_ind = min(df.index)
        so_table['XCol'] = [skipcols_front + 1 + (i // max_rows) for i in so_table.index]  # quotient
        so_table['XCol'] = so_table['XCol'].apply(lambda x: self.get_excel_col_from_int(x))
        so_table['XCell'] = so_table['XCol'] + so_table['XRow'].astype('str')
        # return so format table
        return so_table.reset_index(drop='True')

    def get_cell_colour_col(self, so_table, original_sheet, country_col=None):
        # load parameters if not specified
        if country_col is None:
            country_col = self.__parameters['names']['country_col']
        # run analysis
        so_table['Cell_Colour'] = so_table['XCell'].apply(lambda x: str(original_sheet[x].fill.start_color.index))
        so_table['Cell_Colour'] = [
            '00000000' if colour == '0' or country == 'TOTAL' else colour
            for colour, country in zip(so_table['Cell_Colour'], so_table[country_col])
        ]
        return so_table