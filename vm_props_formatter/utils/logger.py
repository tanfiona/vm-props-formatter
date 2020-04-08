import logging
import logging.handlers
import pandas as pd
import getpass


def format_logs(app_name, is_format=True):
    if is_format:
        # Set up format for saving logs into txt
        logging.root.handlers = []
        logging.basicConfig(format='%(asctime)s : %(levelname)s : %(message)s',
                            level=logging.DEBUG,
                            filename='./outputs/logs.txt',
                            datefmt='%Y-%m-%d %H:%M:%S')
        # Set up format which is simpler for console use
        console = logging.StreamHandler()
        console.setLevel(logging.ERROR)
        formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(message)s')
        console.setFormatter(formatter)
        logging.getLogger("").addHandler(console)
        # Suppress unnecessary warnings
        logging.getLogger('werkzeug').setLevel(logging.ERROR)
        pd.options.mode.chained_assignment = None
        # Include Open App Info
        logging.debug((str(getpass.getuser()) + ' is launching the ' + app_name + ' app ...'))