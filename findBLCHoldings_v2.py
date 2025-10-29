import sys
import requests
import json
import os
import time
import csv
import re
import datetime
import shutil

from tkinter.filedialog import askopenfilename

import pandas as pd
import numpy as np

sys.path.append(os.path.relpath('config/'))
import secrets_local

from datetime import datetime as datetime_1

def get_oclc_token(client_id, client_secret):
    token_url = "https://oauth.oclc.org/token"
    payload = {
        "grant_type": "client_credentials",
        "scope": "configPlatform context:128807"
    }
    response = requests.post(
        token_url,
        data=payload,
        auth=(client_id, client_secret)
    )
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        print("Failed to get token:", response.text)
        sys.exit(1)

now = datetime_1.now() # current date and time
date_time = now.strftime("%m-%d-%Y %H_%M_%S")

oDir = "./Output"
if not os.path.isdir(oDir) or not os.path.exists(oDir):
    os.makedirs(oDir)

filename = askopenfilename(title="Select input file of OCLC Numbers to search WorldCat for")
oclc_df = pd.read_excel(filename, engine='openpyxl')

#url_prefix = "https://worldcat.org/webservices/catalog/content/libraries/"

#oclc_parameter = "oclcsymbol=%5BBET%2CBXM%2CBOS%2CMBU%2CBZM%2CBOP%2CMBB%2CMBW%2CNED%2CNLA%2CMAS%2CUCW%2CAUM%2CBMU%2CSMU%2CSNS%2CULN%2CWQM%2CNHM%2CMVA%2CWLU%2CWCM%2CTFW%2CTFH%2CTFF%2CTUFTV%2CRIU%2CRIX%2CRIN%2CWEL%2CHRM%2CVTU%2CVTM%2CVTVCM%2CBRL%2CIAILL%2C%5D"
#suffix = "&libtype=1&servicelevel=full&frbrGrouping=off&startLibrary=1&maximumLibraries=100&format=json"


url_prefix = "http://www.worldcat.org/webservices/catalog/content"
#column_list = ["BET", "BXM", "BOS", "MBU", "BZM", "BOP", "MBB", "MBW", "NED", "NLA", "MAS", "UCW", "AUM", "BMU", "SMU", "SNS", "ULN", "WQM", "NHM", "MVA", "WLU", "WCM", "TFW", "TFH", "TFF", "TUFTV", "RIU", "RIX", "RIN", "WEL", "HRM", "VTU", "VTM", "VTVCM", "BRL", "IAILL"]

oclc_df = pd.concat([oclc_df, pd.DataFrame(columns=column_list)])
pd.set_option('display.max_columns', None)

print(oclc_df)

# Get OAuth token using credentials from secrets_local.py
access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)

headers = {
    "accept": "application/json",
    "Authorization": f"Bearer {access_token}"
}

for x in range(0, len(oclc_df)):
    oclc_no = str(oclc_df.iloc[x]['Verified OclcNo'])
    url = f"{url_prefix}{oclc_no}?{oclc_parameter}{suffix}"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        try:
            results = response.json()
        except Exception as e:
            print(f"JSON decode error for OCLC {oclc_no}: {e}")
            print("Response text:", response.text)
            continue
    else:
        print(f"HTTP error {response.status_code} for OCLC {oclc_no}")
        print("Response text:", response.text)
        continue

    for library in results.get("library", []):
        oclc_df.at[x, library.get("oclcSymbol", "")] = "x"

pd.set_option('display.max_columns', None)
print(oclc_df)

output_filename = f"{oDir}/Copy of Proposed CDL Corpus Aug 2022 - Output.xlsx - With BLC Holdings {date_time}.xlsx"
oclc_df.to_excel(output_filename, engine="openpyxl", index=False)