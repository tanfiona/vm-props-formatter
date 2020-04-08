import os


def check_create_directory(filename):
    """
    Check and create the directory if it does not exist
    
    Parameters
    ----------
    filename : new filename

    Returns
    -------
    None
    """
    directory = os.path.dirname(filename)
    if not os.path.exists(directory):
        os.makedirs(directory)
