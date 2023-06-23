#!/usr/bin/env python3
"""Pull copy of Total Rec Boat Sales

    Will download smartsheet report as a xlsx file. Allows for override of
    destination file name. On error an email will be sent to the administrator
"""

import os
import sys
from datetime import datetime
from pathlib import Path
from dotenv import load_dotenv
from envelopes import Envelope
from smtplib import SMTPException # allow for silent fail in try exception
import click
import smartsheet  # type: ignore
import traceback


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller

    Arguments:
        relative_path: str -- releative path from current directory + file name

    Returns:
        str -- absoulte path to file
    """
    try:
        # pylint: disable=protected-access
        base_path: str = sys._MEIPASS  # type: ignore
    except AttributeError:
        base_path: str = os.path.abspath(".")  #type: ignore

    return os.path.join(base_path, relative_path)

def split_address(email_address):
    """Return a tuple of (address, name), name may be an empty string
       Can convert the following forms
         exaple@example.com
         <example@exmaple.con>
         Example <example@example.com>
         Example<example@example.com>
    """
    address = email_address.split('<')
    if len(address) == 1:
        return (address[0], '')
    if address[0]:
        return (address[1][:-1], address[0].strip())
    return (address[1][:-1], '')

def mail_results(subject, body, attachment=''):
    """ Send emial with html formatted body and parameters from env"""
    envelope = Envelope(
        from_addr=split_address(os.environ.get('MAIL_FROM')),
        subject=subject,
        html_body=body
    )

    # add standard recepients
    tos = os.environ.get('MAIL_TO','').split(',')
    if tos[0]:
        for to in tos:
            envelope.add_to_addr(to)

    # add carbon coppies
    ccs = os.environ.get('MAIL_CC','').split(',')
    if ccs[0]:
        for cc in ccs:
            envelope.add_cc_addr(cc)

    # add blind carbon copies recepients
    bccs = os.environ.get('MAIL_BCC','').split(',')
    if bccs[0]:
        for bcc in bccs:
            envelope.add_bcc_addr(bcc)

    if attachment:
        envelope.add_attachment(attachment)

    # send the envelope using an ad-hoc connection...
    try:
        _ = envelope.send(
            os.environ.get('MAIL_SERVER'),
            port=os.environ.get('MAIL_PORT'),
            login=os.environ.get('MAIL_LOGIN'),
            password='zcrkyqvgbxkxnjdg',
            tls=True
        )
    except SMTPException:
        print("SMTP EMail error")


@click.command()
@click.option(
    '-f', '--file',
    'input_file',
    required=False,
    help='full path/filemae/extension',
)
@click.option('-v', '--verbose', count=True)
def main(input_file:str, verbose: int)-> None:
    """download smarstsheet report and save as .xlsx file

    Arguments:
        input_file: str -- folder / filename to save spreadsheet as can
                           override .xlsx extension. If no extension is
                           provided then .xlsx will automatically be appended.
        verboste: int   -- enable verbose output mode
    """
    try:
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

        if verbose:
            click.echo(f"Saving spreadsheet as: {path}")
        smart.Reports.get_report_as_excel(report_id, dir_name, file_name)
    except Exception:
        # send message
        mail_results(
            'Smartsheet Total Rec Boat Sales Error',
            '<pre>' + traceback.format_exc() + '</pre>')
        raise
    sys.exit()


if __name__ == "__main__":
    main()  # pylint: disable=no-value-for-parameter
