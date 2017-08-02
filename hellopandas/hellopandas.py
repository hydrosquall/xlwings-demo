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


class AssemblyClient():
    """A class for interacting with the Enigma Assembly API

        Public Documentation:
        https://docs.public.enigma.com

        Todo: add cache:
        https://pypi.python.org/pypi/requests-cache
    """

    def __init__(self, api_key=None):

        if not api_key:
            api_key = os.environ.get("ENIGMA_API_KEY")
        self._api_key = api_key
        headers = {
            "Authorization": "Bearer {0}".format(api_key),
            "User-Agent": 'enigma-assembly-excelwings'
        }

        session = requests.Session()
        session.headers.update(headers)
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


def hello_xlwings():
    """Default function produced by excelwings quickstart

        xlwings functions on Mac are not allowed to take any arguments.
    """
    load_dotenv(find_dotenv())  # Load API keys
    assembly_client = AssemblyClient()
    wb = xw.Book.caller()
    config_sheet = xw.sheets[1]

    # Load config variables
    dataset = config_sheet.range("B1").value
    if dataset:
        dataset_metadata = assembly_client.get_dataset_metadata(dataset)
        snapshot = dataset_metadata['current_snapshot']['id']
        config_sheet.range("C1").value = dataset_metadata['display_name']
    else:
        snapshot = config_sheet.range("B2").value

    # If debugging use this hardcoded snapshot
    # snapshot = "f86381d5-2e8d-4db3-9bcd-77e7f8d87ff0"

    # Get the CSV saved locally
    raw_data = assembly_client.get_snapshot_export(snapshot)
    df = pd.read_csv(io.StringIO(raw_data.decode('utf-8')))

    # Write results to the excel sheet
    # Optional: you could also output the top 900,000 items to provide partial
    # info...
    df.index.name = "PandasIndex"
    height = df.shape[0]
    if height > 900000:  # More conservative than the hard limit, for safety
        output = "{0} has {1} rows: don't open this data in Excel."\
            .format(dataset, height)
    else:
        output = df

    wb.sheets[0].range("A1").value = output
