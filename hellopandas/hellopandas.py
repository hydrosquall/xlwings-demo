"""
hellopandas.py
Cameron Yick
08/02/17

A "hello world" example of how to use Python to connect to the Enigma API.

This module is coupled to the hellopandas.xlsm workbook.
The first sheet in the notebook will contain table from Enigma
The second sheet is for configuring the pandas plugin.

Changing the order of the sheets will break this plugin's functionality.
"""
import io
import os

import pandas as pd
import requests
import xlwings as xw
from dotenv import load_dotenv, find_dotenv
from cachecontrol import CacheControl

# Todo: Investigate alternate lib if necessary
# try https://statcompute.wordpress.com/tag/pysqldf/
from pandasql import sqldf

EXCEL_ROW_LIMIT = 200000


class AssemblyClient():
    """A class for interacting with the Enigma Assembly API

        Public Documentation:
        https://docs.public.enigma.com

        Todo: add cache:
        https://pypi.python.org/pypi/requests-cache
    """

    def __init__(self, api_key=None, use_cache=True):

        if not api_key:
            api_key = os.environ.get("ENIGMA_API_KEY")
        self._api_key = api_key
        headers = {
            "Authorization": "Bearer {0}".format(api_key),
            "User-Agent": 'enigma-assembly-excelwings'
        }

        session = requests.Session()
        session.headers.update(headers)

        if use_cache:
            session = CacheControl(session)

        self._session = session
        self._base_url = "https://public.enigma.com/api"

    def get_dataset_metadata(self, dataset_id):
        url = "datasets/{0}?row_limit=0".format(dataset_id)
        response = self.get(url)
        return response.json()

    def get_snapshot_export(self, snapshot_id):
        url = "export/{0}".format(snapshot_id)
        response = self.get(url)
        return response.content

    def get(self, url):
        url = "{0}/{1}".format(self._base_url, url)
        response = self._session.get(url)
        response.raise_for_status()
        return response


def pysqldf(query, context):
    """Recommended design pattern by docs authors"""
    return sqldf(query, context)


def hello_xlwings():
    """Default function produced by excelwings quickstart

        xlwings functions on Mac are not allowed to take any arguments.
    """

    # Setup
    load_dotenv(find_dotenv())  # Load API keys
    assembly_client = AssemblyClient(use_cache=True)
    wb = xw.Book.caller()
    data_sheet = wb.sheets[0]
    config_sheet = wb.sheets[1]
    preview_sheet = wb.sheets[2]

    # Load config variables
    dataset_id = config_sheet.range("B1").value
    if dataset_id:
        dataset_metadata = assembly_client.get_dataset_metadata(dataset_id)
        snapshot_id = dataset_metadata['current_snapshot']['id']
        config_sheet.range("C1").value = dataset_metadata['display_name']
    else:
        snapshot_id = config_sheet.range("B2").value

    # If no query is provided, by default display some data
    query = config_sheet.range("B3").value
    if not query:
        query = """SELECT * FROM dataset LIMIT {};""".format(EXCEL_ROW_LIMIT)

    # Get the CSV saved locally
    raw_data = assembly_client.get_snapshot_export(snapshot_id)
    dataset = pd.read_csv(io.StringIO(raw_data.decode('utf-8')))

    # Show available columns in a preview pane
    preview_sheet.range('A1').value = dataset.head(10)
    filtered = pysqldf(query, locals())
    data_sheet.clear_contents()

    # Write results to the excel sheet
    height = filtered.shape[0]
    if height > EXCEL_ROW_LIMIT:
        outputCell = "A2"
        data_sheet.range("A1").value = \
            """{0} has {2} rows: we just display the top {1}"
                Consider providing a custom SQL command"""\
                .format(dataset_id, EXCEL_ROW_LIMIT, dataset.shape[0])
        filtered = filtered.head(EXCEL_ROW_LIMIT)
    else:
        outputCell = "A1"

    data_sheet.range(outputCell).value = filtered
