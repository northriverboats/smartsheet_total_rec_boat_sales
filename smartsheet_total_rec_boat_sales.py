#!/usr/bin/env python3
"""Pull copy of Total Rec Boat Sales

    To Dos:
    * add logging
    * add emailing failures
    * unit tests

"""

import logging
import logging.handlers
import os
import sys
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
import click
import smartsheet  # type: ignore



def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller

    Arguments:
        relative_path str -- releative path from current directory + file name

    Returns:
        str -- absoulte path to file
    """
    try:
        # pylint: disable=protected-access
        base_path: str = sys._MEIPASS  # type: ignore
    except AttributeError:
        base_path: str = os.path.abspath(".")  #type: ignore

    return os.path.join(base_path, relative_path)


# ==================== ENALBE LOGGING
# DEBUG + = to stdout
# CRITICAL + = to email
MAIL_SERVER: str = str(os.environ.get("MAIL_SERVER", ''))
MAIL_FROM: str = str(os.environ.get("MAIL_FROM", ''))
MAIL_TO: str = str(os.environ.get("MAIL_TO", ''))
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s')

consoleHandler = logging.StreamHandler(sys.stdout)
consoleHandler.setLevel(logging.DEBUG)
consoleHandler.setFormatter(formatter)

smtpHandler = logging.handlers.SMTPHandler(
    mailhost=MAIL_SERVER,
    fromaddr=MAIL_FROM,
    toaddrs=MAIL_TO,
    subject="alert!"
)
smtpHandler.setLevel(logging.CRITICAL)
smtpHandler.setFormatter(formatter)
logger.addHandler(consoleHandler)
logger.addHandler(smtpHandler)





@click.command()
@click.option(
    '-f', '--file',
    'input_file',
    required=False,
    help='full path/filemae/extension',
)
def main(input_file:str)-> None:
    """download smarstsheet and save to correct folder"""
    load_dotenv(dotenv_path=resource_path(".env"))

    api: str = os.environ.get('SMARTSHEET_API', '')
    report_id: str = os.environ.get('SMARTSHEET_ID', '')
    sheet_name: str = os.environ.get('SMARTSHEET_NAME', '')

    smart = smartsheet.Smartsheet(api)
    smart.assume_user(os.environ.get('SMARTSHEET_USER', ''))

    dated_name: str = datetime.now().strftime(sheet_name)
    full_path: str = os.environ.get('TARGET_DIR','') + '/' + dated_name

    if input_file:
        full_path = input_file

    path: Path = Path(full_path)
    dir_name = str(path.parent.absolute())
    stem: str = path.stem
    file_name: str = path.name
    if stem == file_name:
        file_name += '.xlsx'

    smart.Reports.get_report_as_excel(report_id, dir_name, file_name)
    sys.exit()


if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter
