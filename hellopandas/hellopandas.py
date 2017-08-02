import xlwings as xw
import pandas as pd
import requests
import io
import os
from dotenv import load_dotenv, find_dotenv


def get_assembly_session(api_key=None):
    """
    """

    if not api_key:
        api_key = os.environ.get("ENIGMA_API_KEY")
    headers = {
        "Authorization": "Bearer {0}".format(api_key)
    }

    session = requests.Session()
    session.headers.update(headers)
    return session


def hello_xlwings():
    load_dotenv(find_dotenv())  # Load API keys
    session = get_assembly_session()
    wb = xw.Book.caller()
    config_sheet = xw.sheets[1]
    base_url = "https://public.enigma.com/api"

    # Load config variables
    dataset = config_sheet.range("B1").value
    if dataset:
        dataset_url = "{0}/datasets/{1}?row_limit=0".format(base_url, dataset)
        response = session.get(dataset_url)
        response.raise_for_status()
        dataset_metadata = response.json()
        snapshot = dataset_metadata['current_snapshot']['id']
        config_sheet.range("C1").value = dataset_metadata['display_name']
    else:
        snapshot = config_sheet.range("B2").value

    # If debugging use this hardcoded snapshot
    # snapshot = "f86381d5-2e8d-4db3-9bcd-77e7f8d87ff0"
    snap_url = "{0}/export/{1}".format(base_url, snapshot)
    response = session.get(snap_url)
    response.raise_for_status()
    raw_data = response.content

    # df = pd.DataFrame([[1.1, 2.2], [3.3, None]], columns=['one', 'two'])
    df = pd.read_csv(io.StringIO(raw_data.decode('utf-8')))
    df.index.name = "Row Number"
    wb.sheets[0].range("A1").value = df[df.columns]
