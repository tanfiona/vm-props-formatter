import os


def check_create_directory(file_path):
    """
    Check and create the directory if it does not exist
    
    Parameters
    ----------
    file_path : new file path

    Returns
    -------
    None
    """
    directory = os.path.dirname(file_path)
    if not os.path.exists(directory):
        os.makedirs(directory)
