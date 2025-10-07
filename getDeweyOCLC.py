#!/usr/bin/env python3
import pandas as pd
import json
import re
import requests
import time
import os
from tkinter import filedialog

import sys
sys.path.append('config/')
import secrets_local

# --- Global dicts to store OCLC results ---
title_dict = {}
dewey_dict = {}
date_dict = {}
grouped_oclc_df = None
access_token = ""

# -------------------------------
# Auth
# -------------------------------
def get_oclc_token(client_id, client_secret):
    token_url = "https://oauth.oclc.org/token"
    payload = {"grant_type": "client_credentials", "scope": "wcapi"}

    response = requests.post(token_url, data=payload, auth=(client_id, client_secret))
    if response.status_code == 200:
        return response.json()["access_token"]
    else:
        raise Exception(f"Failed to obtain token: {response.status_code} - {response.text}")

# -------------------------------
# OCLC Search
# -------------------------------
def search_oclc(oclc):
    global access_token, title_dict, dewey_dict, date_dict

    print(f"OCLC Number to search: ||{oclc}||")
    url_prefix = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs?q=no:"

    if not access_token:
        access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)

    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    max_retries = 5
    backoff = 5

    for attempt in range(max_retries):
        try:
            resp = requests.get(
                url_prefix + str(oclc) + "&rows=1",
                headers=headers,
                timeout=15
            )

            # --- Handle expired token ---
            if resp.status_code == 401:
                print("Access token expired, refreshing...")
                access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)
                headers["Authorization"] = f"Bearer {access_token}"
                continue  # retry immediately

            if resp.status_code == 200:
                data = resp.json()

                if data.get("numberOfRecords", 0) > 0:
                    rec = data["bibRecords"][0]

                    titleAuthor = rec.get("title", {}).get("mainTitles", [{}])[0].get("text", "No Title Found")
                    dewey = rec.get("classification", {}).get("dewey", "No Dewey Found")
                    publicationDate = rec.get("date", {}).get("publicationDate", "No Publication Date Found")

                    title_dict[oclc] = titleAuthor
                    dewey_dict[oclc] = dewey
                    date_dict[oclc] = publicationDate
                    return  # success
                else:
                    print(f"Attempt {attempt+1}: No records for OCLC {oclc}")

            else:
                print(f"Attempt {attempt+1}: HTTP {resp.status_code} - {resp.text}")

        except requests.exceptions.Timeout:
            print(f"Attempt {attempt+1}: timeout for OCLC {oclc}")

        except Exception as e:
            print(f"Attempt {attempt+1}: error {e}")

        time.sleep(backoff)
        backoff *= 2

    # fallback if all retries fail
    title_dict[oclc] = "No OCLC Results - not found"
    dewey_dict[oclc] = "No OCLC Results - not found"
    date_dict[oclc] = "No OCLC Results - not found"

# -------------------------------
# Spreadsheet Processing
# -------------------------------
def process_spreadsheet(oclc_list, output_csv_path=None, chunk_sleep=0.2):
    global grouped_oclc_df, access_token

    if output_csv_path is None:
        from datetime import datetime
        timestamp_filename = datetime.now().strftime("%Y%m%d%H%M%S")
        output_csv_path = f"Output/Library Statistics with Merged OCLC Data {timestamp_filename}.csv"
    os.makedirs(os.path.dirname(output_csv_path), exist_ok=True)

    # Ensure token present
    if not access_token:
        access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)

    grouped_oclc_df["_OCLC_norm_"] = grouped_oclc_df["OCLC Control Number (035a)"].astype(str).str.strip()

    if os.path.exists(output_csv_path):
        os.remove(output_csv_path)

    header_written = False
    total = len(oclc_list)

    for idx, raw_oclc in enumerate(oclc_list, start=1):
        oclc_str = str(raw_oclc).strip()
        if not oclc_str or oclc_str.lower().startswith("nan"):
            continue

        search_oclc(oclc_str)

        matched = grouped_oclc_df[grouped_oclc_df["_OCLC_norm_"] == oclc_str]
        if not matched.empty:
            matched = matched.copy()
            key_col = "OCLC Control Number (035a)"
            matched["OCLC TitleAuthor"] = matched[key_col].apply(lambda k: title_dict.get(str(k).strip(), "No OCLC Results - not found"))
            matched["OCLC Dewey"] = matched[key_col].apply(lambda k: dewey_dict.get(str(k).strip(), "No OCLC Results - not found"))
            matched["OCLC Publication Date"] = matched[key_col].apply(lambda k: date_dict.get(str(k).strip(), "No OCLC Results - not found"))
            matched["OCLC Number"] = oclc_str

            matched.to_csv(output_csv_path, mode="a", index=False, header=not header_written)
            header_written = True
            print(f"[{idx}/{total}] Appended {len(matched)} row(s) for OCLC {oclc_str}")

        time.sleep(chunk_sleep)

    if "_OCLC_norm_" in grouped_oclc_df.columns:
        grouped_oclc_df.drop(columns=["_OCLC_norm_"], inplace=True)

    print(f"Incremental output complete: {output_csv_path}")

# -------------------------------
# Main
# -------------------------------
def main():
    global grouped_oclc_df
    start_time = time.perf_counter()

    print("Merging course, department, and SUSHI data; retrieving OCLC info.\n")
    oclc_answer = input("Do you already have OCLC data (yes/no)?: ").lower()
    oclc_bool = oclc_answer in ("yes", "y")

    output_dir = os.path.join(os.getcwd(), "Output")
    os.makedirs(output_dir, exist_ok=True)
    input_dir = os.path.join(os.getcwd(), "input")
    os.makedirs(input_dir, exist_ok=True)

    semester_counts_file_path = filedialog.askopenfilename(title="Upload Excel file of semester course counts")
    academic_deparment_filepath = filedialog.askopenfilename(title="Upload Excel file of academic departments")
    sushi_counts_file_path = filedialog.askopenfilename(title="Upload Excel file of SUSHI ebook counts")
    oclc_data_file = filedialog.askopenfilename(title="Upload Excel file of OCLC data") if oclc_bool else ""

    semester_df = pd.read_excel(semester_counts_file_path, dtype={'ISBN (Normalized)': "str"}, engine="openpyxl")
    academic_df = pd.read_excel(academic_deparment_filepath, engine="openpyxl")
    sushi_df = pd.read_excel(sushi_counts_file_path, dtype={'Normalized ISBN': 'str'}, engine="openpyxl")

    semester_df['Associated Course Code'] = semester_df['Associated Course Code'].str.split(';')
    semester_df = semester_df.explode('Associated Course Code').reset_index(drop=True)

    merged_df = pd.merge(semester_df, academic_df, how='right',
                         left_on=['MMS Id', 'Associated Course Code'],
                         right_on=['MMS Id', 'Course Code']).reset_index(drop=True)

    merged_df['ISBN (Normalized)'] = merged_df['ISBN (Normalized)_x'].combine_first(merged_df['ISBN (Normalized)_y'])
    grouped_df = merged_df.groupby(
        ["Library Name", "Location Name", "Item Policy", "Loan Fiscal Year",
         "ISBN (Normalized)_x", "ISBN (Normalized)_y", "ISBN (Normalized)",
         "OCLC Control Number (035a)_x", "OCLC Control Number (035a)_y", "MMS Id",
         "Loans (Not In House)", "Course Fiscal Year"],
        as_index=False, dropna=False
    ).agg({
        'Course Code': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
        'Course Name': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
        'Associated Course Code': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
        'Academic Department Description': lambda x: '; '.join(sorted(set(x.dropna().astype(str)))),
    })

    grouped_df['ISBN (Normalized)'] = grouped_df['ISBN (Normalized)'].str.split(';')
    exploded_df = grouped_df.explode('ISBN (Normalized)').reset_index(drop=True)

    exploded_df['ISBN (Normalized)'] = exploded_df['ISBN (Normalized)'].apply(
        lambda x: re.sub(r'[\D-]', '', str(x)) if pd.notna(x) else ""
    )

    merged_with_sushi_df = pd.merge(exploded_df, sushi_df, how='left',
                                    left_on=['ISBN (Normalized)', 'Course Fiscal Year'],
                                    right_on=['Normalized ISBN', 'Usage Date Fiscal Year'])

    merged_with_sushi_df["OCLC Control Number (035a)"] = merged_with_sushi_df['OCLC Control Number (035a)_x'].combine_first(
        merged_with_sushi_df['OCLC Control Number (035a)_y']
    )
    merged_with_sushi_df = merged_with_sushi_df.drop(columns=[
        'OCLC Control Number (035a)_x', 'OCLC Control Number (035a)_y',
        'ISBN (Normalized)_x', 'ISBN (Normalized)_y'
    ])

    grouped_oclc_df = merged_with_sushi_df.groupby(
        ["Library Name", "Location Name", "Item Policy", "Loan Fiscal Year",
         "OCLC Control Number (035a)", "MMS Id", "Loans (Not In House)",
         "Course Fiscal Year", "Course Code", "Course Name", "Associated Course Code",
         "Academic Department Description", "Platform", "Normalized Title",
         "Display Title", "Normalized ISBN", "Title Identifier Count",
         "Usage Measures Total", "Section Type", "Usage Date Fiscal Year"],
        as_index=False, dropna=False
    ).agg({
        "ISBN (Normalized)": lambda x: '; '.join(sorted(set(x.dropna().astype(str))))
    })

    grouped_oclc_df['OCLC Control Number (035a)'] = grouped_oclc_df['OCLC Control Number (035a)'].astype(str).str.strip()

    unique_oclcs = grouped_oclc_df['OCLC Control Number (035a)'].unique()

    from datetime import datetime
    timestamp_filename = datetime.now().strftime("%Y%m%d%H%M%S")
    output_filename = f"Output/Library Statistics with Merged OCLC Data {timestamp_filename}.csv"

    if oclc_bool:
        print("Merging with provided OCLC data file.")
        oclc_df = pd.read_csv(oclc_data_file, dtype={'OCLC Control Number (035a)': "str"})
        oclc_df['OCLC Control Number (035a)'] = oclc_df['OCLC Control Number (035a)'].astype(str).str.strip()
        grouped_oclc_df = pd.merge(grouped_oclc_df, oclc_df, how='left',
                                   on=['OCLC Control Number (035a)', 'Course Code'])
        grouped_oclc_df.to_csv(output_filename, index=False)
        print(f"Wrote merged dataframe to: {output_filename}")
    else:
        process_spreadsheet(unique_oclcs, output_csv_path=output_filename)
        print(f"Incremental output written to: {output_filename}")

if __name__ == "__main__":
    main()
