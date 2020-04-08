import json
import os
from vm_props_formatter.utils.file_organizer import check_create_directory

def read_json(filename):
    """
    Read the data from a JSON file

    Parameters
    ----------
    filename : str
        JSON filename

    Returns
    -------
    data : dict
        Dictionary of values from the JSON file
    """
    return json.load(open(filename, 'r', encoding='utf-8')) if os.path.isfile(filename) else None

def write_json(data, filename):
    """
    Write the data into a JSON file
    
    Parameters
    ----------
    data : dict
        Dictionary of values
    filename : str
        JSON filename
    
    Returns
    -------
    None
    """
    check_create_directory(filename)
    file = open(filename, 'w')
    file.write(json.dumps(data))
    file.close()
