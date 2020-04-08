import copy
import glob
import os
from utils.file_organizer import check_create_directory

class StoreConsolidationManager(object):
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
            "country_col": "COUNTRY NAME",
            "store_col": "STORE NAME",
            "main_cols": ["COUNTRY NAME", "STORE NAME"],
            "entity_list": ["CKS","CKI","CKC"]
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