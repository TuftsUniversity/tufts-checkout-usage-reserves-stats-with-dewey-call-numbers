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


title_dict = {}
dewey_dict = {}
date_dict = {}  
oclc_number_dict = {}
grouped_oclc_df = None
access_token = ""

def read_excel_file(file_path):
    df = pd.read_excel(file_path, skiprows=0, header=1, engine="openpyxl", dtype={'Course': 'str', 'Dept': 'str', 'Sec': 'str', 'EAN-13': 'str'})
    df = df.fillna("")


    return df


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

def search_oclc(oclc):
    
    global title_dict, dewey_dict, date_dict, grouped_oclc_df, oclc_number_dict, access_token
    print("OCLC Number  to search: ||" + oclc + "||")

    url_prefix = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs?q=no:"


    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    

# ...existing code...
def search_oclc(oclc):
    """
    Query OCLC for a single OCLC number and return a dict:
    { 'title': str, 'dewey': str, 'pub_date': str, 'oclc_number': str }
    Uses global access_token (must be set before calling).
    """
    global access_token

    oclc_str = str(oclc).strip()
    if oclc_str == "" or oclc_str.lower().startswith("nan"):
        return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}

    url_prefix = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs?q=no:"
    url = f"{url_prefix}{urllib.parse.quote(oclc_str)}&rows=1"

    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    max_retries = 5
    backoff = 5
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                try:
                    data = resp.json()
                except Exception:
                    return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}

                # Safely extract common fields (fall back to empty strings)
                title = ""
                dewey = ""
                pub_date = ""

                try:
                    br = data.get("bibRecords") or []
                    if isinstance(br, list) and len(br) > 0:
                        rec = br[0]
                        # title
                        title = (
                            rec.get("title", {})
                               .get("mainTitles", [{}])[0]
                               .get("text", "") if rec.get("title") else ""
                        ) or title
                        # dewey
                        if rec.get("classification"):
                            cls = rec["classification"]
                            if isinstance(cls, dict) and "dewey" in cls:
                                dewey = cls.get("dewey") or dewey
                            elif isinstance(cls, list) and len(cls) > 0 and isinstance(cls[0], dict):
                                dewey = cls[0].get("dewey") or dewey
                        # publication date
                        if rec.get("date") and isinstance(rec["date"], dict):
                            pub_date = rec["date"].get("publicationDate") or pub_date
                except Exception:
                    pass

                return {
                    "title": "" if title is None else str(title),
                    "dewey": "" if dewey is None else str(dewey),
                    "pub_date": "" if pub_date is None else str(pub_date),
                    "oclc_number": oclc_str
                }

            elif resp.status_code in (429, 503):
                # rate limited / service unavailable -> backoff and retry
                time.sleep(backoff)
                backoff *= 2
                continue
            else:
                # other HTTP errors -> return empties (no retry)
                return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}

        except requests.exceptions.Timeout:
            time.sleep(backoff)
            backoff *= 2
            continue
        except Exception:
            # unexpected error; return empties
            return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}

    # exhausted retries
    return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}


def process_spreadsheet(oclc_list, output_csv_path=None, chunk_sleep=0.2):
    """
    Iterate unique OCLC numbers, query OCLC, and append matching rows from grouped_oclc_df
    to the CSV incrementally. Creates header once.
    """
    global grouped_oclc_df, access_token

    if output_csv_path is None:
        from datetime import datetime
        timestamp_filename = datetime.now().strftime("%Y%m%d%H%M%S")
        output_csv_path = f"Output/Library Statistics with Merged OCLC Data {timestamp_filename}.csv"
    os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

    # ensure token present
    access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)

    # normalize key used for matching
    key_col = "OCLC Control Number (035a)"
    grouped_oclc_df["_OCLC_norm_"] = grouped_oclc_df[key_col].astype(str).str.strip()

    # remove existing file to start fresh
    if os.path.exists(output_csv_path):
        os.remove(output_csv_path)

    header_written = False
    total = len(oclc_list)
    for idx, raw_oclc in enumerate(oclc_list, start=1):
        oclc_str = str(raw_oclc).strip()
        if oclc_str == "" or oclc_str.lower().startswith("nan"):
            # still append any rows that have blank OCLC if present
            matched = grouped_oclc_df[grouped_oclc_df["_OCLC_norm_"] == oclc_str]
            if not matched.empty:
                matched = matched.copy()
                matched["OCLC TitleAuthor"] = ""
                matched["OCLC Dewey"] = ""
                matched["OCLC Publication Date"] = ""
                matched["OCLC Number"] = oclc_str
                matched.to_csv(output_csv_path, mode="a", index=False, header=not header_written)
                header_written = True
                print(f"[{idx}/{total}] Appended {len(matched)} row(s) for blank OCLC")
            continue

        # get OCLC info
        info = search_oclc(oclc_str)

        # find matching rows in grouped_oclc_df
        matched = grouped_oclc_df[grouped_oclc_df["_OCLC_norm_"] == oclc_str]
        if matched.empty:
            print(f"[{idx}/{total}] OCLC {oclc_str}: no matching rows")
        else:
            matched = matched.copy()
            matched["OCLC TitleAuthor"] = info.get("title", "")
            matched["OCLC Dewey"] = info.get("dewey", "")
            matched["OCLC Publication Date"] = info.get("pub_date", "")
            matched["OCLC Number"] = info.get("oclc_number", oclc_str)

            # append to CSV; write header only once
            matched.to_csv(output_csv_path, mode="a", index=False, header=not header_written)
            header_written = True
            print(f"[{idx}/{total}] Appended {len(matched)} row(s) for OCLC {oclc_str}")

        # throttle
        time.sleep(chunk_sleep)

    # cleanup
    if "_OCLC_norm_" in grouped_oclc_df.columns:
        grouped_oclc_df.drop(columns=["_OCLC_norm_"], inplace=True)

    print(f"Incremental output complete: {output_csv_path}")
# ...existing code...
def main():

    import time

    global grouped_oclc_df
    start_time = time.perf_counter()

    print("This script merges semester course data with academic department data and SUSHI ebook usage data, then retrieves OCLC information based on OCLC Control Numbers.\n")
    oclc_answer = input("Do you already have OCLC data (yes/no)?: ")

    oclc_bool = False
    if oclc_answer.lower() == "yes" or oclc_answer.lower() == "y":
        oclc_bool = True
    output_dir = os.path.join(os.getcwd(), "Output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    input_dir = os.path.join(os.getcwd(), "input")
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
    semester_counts_file_path = filedialog.askopenfilename(title="Upload Excel file of semester course counts")

    academic_deparment_filepath = filedialog.askopenfilename(title="Upload Excel file of academic departments")
    
    sushi_counts_file_path = filedialog.askopenfilename(title="Upload Excel file of SUSHI ebook counts")# = semester_counts_file_path.replace("\\", "/")
    if oclc_bool:
        oclc_data_file = filedialog.askopenfilename(title="Upload Excel file of OCLC data") if oclc_answer.lower() == "yes" or oclc_answer.lower() == "y" else ""
    semester_df = pd.read_excel(semester_counts_file_path, skiprows=0, header=0, dtype={'ISBN (Normalized)': "str", "Associated Course Code": str, "Associated Course Name": str}, engine="openpyxl")
    print("Length of semester df")
    print(len(semester_df))
    academic_df = pd.read_excel(academic_deparment_filepath, skiprows=0, header=0, engine="openpyxl")
    print("Length of academic df")
    print(len(academic_df))
    sushi_df = pd.read_excel(sushi_counts_file_path, skiprows=0, header=0, dtype={'Normalized ISBN': 'str'}, engine="openpyxl")   
    semester_df['Associated Course Code'] = semester_df['Associated Course Code'].str.split(';')
    semester_df = semester_df.explode('Associated Course Code').reset_index(drop=True)
    # academic_df = academic_df.drop(columns=['ISBN'])
    # semester_df['ISBN (Normalized)'] = semester_df['ISBN (Normalized)'].str.split(';')



    merged_df = pd.merge(semester_df, academic_df, how='right', left_on=['MMS Id', 'Associated Course Code'], right_on=['MMS Id', 'Course Code'])
    #merged_df = pd.merge(semester_df, academic_df, how='left', left_on=['MMS Id', 'Course Code'], right_on=['MMS Id', 'Course Code'])
    merged_df = merged_df.reset_index(drop=True)  
    merged_df.to_csv("Output/Merged Semester Loans and Course Subjects Unrolled.csv", index=False)
    #sys.exit()
    # Example: merge if left['A'] OR left['B'] matches right['X']

    merged_df['ISBN (Normalized)'] = merged_df['ISBN (Normalized)_x'].combine_first(merged_df['ISBN (Normalized)_y'])   

    # This rolls back titles such that course information is grouped by fiscal year
    # ISBN remains separate because separate ISBNs for each title have alreay been exploded out
    grouped_df = merged_df.groupby(
    ["Library Name", "Location Name", "Item Policy", "Loan Fiscal Year", "ISBN (Normalized)_x", "ISBN (Normalized)_y", "ISBN (Normalized)", "OCLC Control Number (035a)_x", "OCLC Control Number (035a)_y", "MMS Id", "Loans (Not In House)", "Course Fiscal Year"],
    as_index=False,
    dropna=False
    ).agg({
        'Course Code': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
        'Course Name': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
        'Associated Course Code': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
        'Academic Department Description': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),

    })

    #grouped_df = grouped_df.apply(lambda row: list(set(row['ISBN (Normalized)_x'].split(";") +  row["ISBN (Normalized)_y"].split(";"))))
    pd.set_option('display.max_columns', None)
    print(grouped_df)
        #sys.exit()
    two_time = time.perf_counter()
    elapsed_time = two_time - start_time
    print(f"Elapsed Time: {elapsed_time} seconds")

    grouped_df.to_csv("Output/Merged Semester Loans and Course Subjects.csv", index=False)

    grouped_df['ISBN (Normalized)'] = grouped_df['ISBN (Normalized)'].str.split(';')
    exploded_df = grouped_df.explode('ISBN (Normalized)').reset_index(drop=True)
    
    exploded_df.to_csv("Output/Exploded ISBN Merged Semester Loans and Course Subjects.csv", index=False)
    
    exploded_df['ISBN (Normalized)'] = exploded_df['ISBN (Normalized)'].apply(
        lambda x: re.sub(r'[\D-]', '', str(x)) if pd.notna(x) else ""
    )

    #sys.exit()


    merged_with_sushi_df = pd.merge(exploded_df, sushi_df, how='left', left_on=['ISBN (Normalized)', 'Course Fiscal Year'], right_on=['Normalized ISBN', 'Usage Date Fiscal Year'])


    print(merged_with_sushi_df)
    merged_with_sushi_df.to_csv("Output/Merged Semester Loans and Course Subjects with SUSHI.csv", index=False)
    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    print(f"Elapsed Time: {elapsed_time} seconds")
    #sys.exit()


    merged_with_sushi_df["OCLC Control Number (035a)"] = merged_with_sushi_df['OCLC Control Number (035a)_x'].combine_first(merged_with_sushi_df['OCLC Control Number (035a)_y']) 
    merged_with_sushi_df = merged_with_sushi_df.drop(columns=['OCLC Control Number (035a)_x', 'OCLC Control Number (035a)_y', 'ISBN (Normalized)_x', 'ISBN (Normalized)_y'])

    merged_with_sushi_df.to_csv("Output/Merged Semester Loans and Course Subjects with SUSHI Cleaned.csv", index=False)
    #merged_with_sushi_df = merged_with_sushi_df.drop(columns=['OCLC Control
    grouped_oclc_df = merged_with_sushi_df.groupby(
    [
        "Library Name", 
        "Location Name", 
        "Item Policy", 
        "Loan Fiscal Year", 
        "OCLC Control Number (035a)", 
        "MMS Id", 
        "Loans (Not In House)", 
        "Course Fiscal Year", 
        "Course Code", 
        "Course Name", 
        "Associated Course Code", 
        "Academic Department Description", 
        "Platform", 
        "Normalized Title", 
        "Display Title", 
        "Normalized ISBN", 
        "Title Identifier Count", 
        "Usage Measures Total", 
        "Section Type", 
        "Usage Date Fiscal Year"

    ],
    as_index=False, 
    dropna=False
    ).agg({
        "ISBN (Normalized)": lambda x: '; '.join(sorted(set(x.dropna().astype(str))))
    })


    print(grouped_oclc_df)

    grouped_oclc_df['OCLC TitleAuthor'] = ""
    grouped_oclc_df['OCLC Dewey'] = ""
    grouped_oclc_df['OCLC Publication Date'] = ""
    grouped_oclc_df.to_csv("Output/Grouped by OCLC Control Number.csv", index=False)
    #sys.exit()

    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    print(f"Elapsed Time: {elapsed_time} seconds")
    grouped_oclc_df['OCLC Control Number (035a)'] = grouped_oclc_df['OCLC Control Number (035a)'].astype(str).str.strip()   
    number_of_mms_ids = len(grouped_oclc_df['MMS Id'].unique())
    number_of_oclcs = len(grouped_oclc_df['OCLC Control Number (035a)'].unique())
    
    unique_oclcs = grouped_oclc_df['OCLC Control Number (035a)'].unique()
    #print("Unique OCLC Control Numbers:")
    print(f"Number of unique MMS IDs: {number_of_mms_ids}")
    print(f"Number of unique OCLC Control Numbers: {number_of_oclcs}")  
    
    # prepare output filename with timestamp and ensure output folder exists
    from datetime import datetime
    now = datetime.now()
    timestamp_filename = now.strftime("%Y%m%d%H%M%S")
    output_filename = f"Output/Library Statistics with Merged OCLC Data {timestamp_filename}.csv"
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)

    if oclc_bool:
        print("You indicated you already have OCLC data. Merging it now.")
        oclc_df = pd.read_csv(oclc_data_file, skiprows=0, header=0, dtype={'OCLC Control Number (035a)': "str"})
        oclc_df = oclc_df[['OCLC Control Number (035a)', 'OCLC TitleAuthor', 'OCLC Dewey', 'OCLC Publication Date', 'Course Code']]
        oclc_df['OCLC Control Number (035a)'] = oclc_df['OCLC Control Number (035a)'].astype(str).str.strip()
        grouped_oclc_df = pd.merge(grouped_oclc_df, oclc_df, how='left', left_on=['OCLC Control Number (035a)', 'Course Code'], right_on=['OCLC Control Number (035a)', 'Course Code'])
        # write merged dataframe once (has OCLC data from supplied file)
        grouped_oclc_df.to_csv(output_filename, index=False)
        print(f"Wrote merged dataframe to: {output_filename}")
    else:
        # incremental processing — pass the timestamped filename so progress is saved there
        process_spreadsheet(unique_oclcs, output_csv_path=output_filename)
        print(f"Incremental output written to: {output_filename}")

    # ...existing code...




if __name__ == "__main__":
    main()
