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

    

def search_oclc(oclc):
    print("OCLC Number to search: ||" + str(oclc) + "||")

    url_prefix = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs?q=no:"

    access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)
    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    max_retries = 5
    backoff = 5  # seconds, will increase exponentially

    for attempt in range(max_retries):
        try:
            resp = requests.get(
                url_prefix + str(oclc) + "&rows=1",
                headers=headers,
                timeout=5
            )

            if resp.status_code == 200:
                data = resp.json()

                titleAuthor = (
                    data['bibRecords'][0]['title']['mainTitles'][0]['text']
                    if 'bibRecords' in data and len(data['bibRecords']) > 0
                    and 'title' in data['bibRecords'][0]
                    and 'mainTitles' in data['bibRecords'][0]['title']
                    and len(data['bibRecords'][0]['title']['mainTitles']) > 0
                    and 'text' in data['bibRecords'][0]['title']['mainTitles'][0]
                    else 'No Title Found'
                )

                dewey = (
                    data['bibRecords'][0]['classification']['dewey']
                    if 'bibRecords' in data and len(data['bibRecords']) > 0
                    and 'classification' in data['bibRecords'][0]
                    and 'dewey' in data['bibRecords'][0]['classification']
                    else 'No Dewey Found'
                )

                publicationDate = (
                    data['bibRecords'][0]['date']['publicationDate']
                    if 'bibRecords' in data and len(data['bibRecords']) > 0
                    and 'date' in data['bibRecords'][0]
                    and 'publicationDate' in data['bibRecords'][0]['date']
                    else 'No Publication Date Found'
                )

                # ✅ Assign into the dicts
                title_dict[oclc] = titleAuthor
                dewey_dict[oclc] = dewey
                date_dict[oclc] = publicationDate

            else:
                print(f"Attempt {attempt+1}: HTTP {resp.status_code} - {resp.text}")
                 # exponential backoff before next retry
                time.sleep(backoff)
                backoff *= 2


        except requests.exceptions.Timeout:
            print(f"Attempt {attempt+1}: timeout for OCLC {oclc}")
             # exponential backoff before next retry
            time.sleep(backoff)
            backoff *= 2


        except Exception as e:
            print(f"Attempt {attempt+1}: error {e}")
             # exponential backoff before next retry
            time.sleep(backoff)
            backoff *= 2


       
    else:
        title_dict[oclc] = "No OCLC Results - not found"
        dewey_dict[oclc] = "No OCLC Results - not found"
        date_dict[oclc] = "No OCLC Results - not found"
        

  
def process_spreadsheet(oclc_list):

   


    global grouped_oclc_df, title_dict, dewey_dict, date_dict, access_token


    access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)


    results = []

    x = 0
    for oclc in oclc_list:
        # if x == 25:
        #     break
        if x != 0 and x % 100 == 0:
            time.sleep(30)
        x += 1

        search_oclc(oclc)

    

    



    

    grouped_oclc_df['OCLC TitleAuthor'] = grouped_oclc_df['OCLC Control Number (035a)'].apply(lambda x: title_dict.get(x))
    grouped_oclc_df['OCLC Dewey'] = grouped_oclc_df['OCLC Control Number (035a)'].apply(lambda x: dewey_dict.get(x))
    grouped_oclc_df['OCLC Publication Date'] = grouped_oclc_df['OCLC Control Number (035a)'].apply(lambda x: date_dict.get(x))  
    grouped_oclc_df['OCLC Number'] = grouped_oclc_df['OCLC Control Number (035a)'].apply(lambda x: oclc_number_dict.get(x))         
    
    pd.set_option('display.max_columns', None)





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
    
    if oclc_bool:
        print("You indicated you already have OCLC data. Merging it now.")
        oclc_df = pd.read_csv(oclc_data_file, skiprows=0, header=0, dtype={'OCLC Control Number (035a)': "str"})
        print("Length of OCLC df")
        print(len(oclc_df))

        oclc_df = oclc_df[['OCLC Control Number (035a)', 'OCLC TitleAuthor', 'OCLC Dewey', 'OCLC Publication Date', 'Course Code']]
        oclc_df['OCLC Control Number (035a)'] = oclc_df['OCLC Control Number (035a)'].astype(str).str.strip()   

        grouped_oclc_df = pd.merge(grouped_oclc_df, oclc_df, how='left', left_on=['OCLC Control Number (035a)', 'Course Code'], right_on=['OCLC Control Number (035a)', 'Course Code'])
        
        print("After merging with existing OCLC data")
        print(grouped_oclc_df)


        
    else:
        process_spreadsheet(unique_oclcs)
        #grouped_oclc_df.to_csv("Output/Library Statistics with OCLC Data.csv", index=False)           


    # grouped_oclc_df = grouped_oclc_df.groupby(
    # [
    #     "Library Name", 
    #     "Location Name", 
    #     "Item Policy", 
    #     "Loan Fiscal Year", 
    #     "ISBN (Normalized)",
    #     "MMS Id", 
    #     "Loans (Not In House)", 
    #     "Course Fiscal Year", 
    #     "Course Code", 
    #     "Course Name", 
    #     "Associated Course Code", 
    #     "Academic Department Description", 
    #     "Platform", 
    #     "Normalized Title", 
    #     "Display Title", 
    #     "Normalized ISBN", 
    #     "Title Identifier Count", 
    #     "Usage Measures Total", 
    #     "Section Type", 
    #     "Usage Date Fiscal Year",
    #     "OCLC TitleAuthor",
    #     "OCLC Dewey",
    #     "OCLC Publication Date",

    # ],
    # as_index=False, 
    # dropna=False
    # ).agg({
    #     "OCLC Control Number (035a)": lambda x: '; '.join(sorted(set(x.dropna().astype(str))))
    # })

    from datetime import datetime
        # Get the current datetime object
    now = datetime.now()

    # Format the datetime object into a filename-safe string
    timestamp_filename = now.strftime("%Y%m%d%H%M%S") 
    grouped_oclc_df.to_csv(f"Output/Library Statistics with Merged OCLC Data {timestamp_filename}.csv", index=False)
     
    



if __name__ == "__main__":
    main()
