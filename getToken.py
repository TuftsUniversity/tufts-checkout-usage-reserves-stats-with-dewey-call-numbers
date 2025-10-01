#!/usr/bin/env python3
import pandas as pd
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import numbers
import sys
sys.path.append('config/')
import secrets_local
#from tkinter.filedialog import askopenfilename
from tkinter import filedialog

import pandas as pd
import json
import re
import requests
import urllib
import xml.etree.ElementTree as ET
import time
import os

def get_oclc_token(client_id, client_secret):
    token_url = "https://oauth.oclc.org/token"
    payload = {
        "grant_type": "client_credentials",
        "scope": "wcapi"
    }

    try:
        response = requests.post(
            token_url,
            data=payload,
            auth=(client_id, client_secret)
        )
        if response.status_code == 200:
            return response.json()["access_token"]
        else:
            raise Exception(f"Failed to obtain token: {response.status_code} - {response.text}")    
    except:
        raise Exception(f"Failed to obtain token: {response.status_code} - {response.text}") 






access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)

print(access_token)