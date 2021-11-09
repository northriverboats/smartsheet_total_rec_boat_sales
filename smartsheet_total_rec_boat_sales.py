#!/usr/bin/env python3
"""pull copy of Total Rec Boat Sales"""

import os
import sys
from datetime import datetime
from dotenv import load_dotenv
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

def main()-> None:
    """download smarstsheet and save to correct folder"""
    load_dotenv(dotenv_path=resource_path(".env"))

    api: str = os.environ.get('SMARTSHEET_API', '')
    report_id: str = os.environ.get('SMARTSHEET_ID', '')
    sheet_name: str = os.environ.get('SMARTSHEET_NAME', '')

    smart = smartsheet.Smartsheet(api)
    smart.assume_user(os.environ.get('SMARTSHEET_USER'))

    file_name: str = datetime.now().strftime(sheet_name)
    smart.Reports.get_report_as_excel(report_id, '/tmp', file_name)
    sys.exit()


if __name__ == "__main__":
    main()
