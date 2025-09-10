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


barcode_array = []

def read_excel_file(file_path):
    df = pd.read_excel(file_path, skiprows=0, header=1, engine="openpyxl", dtype={'Course': 'str', 'Dept': 'str', 'Sec': 'str', 'EAN-13': 'str'})
    df = df.fillna("")


    return df
def parse_xml_response(xml_string, json_course, format_type, row, api_bib, open_library_title, open_library_author):
    results = []

    electronic_record_array = []


    # Parse XML
    #try:



    root = ET.fromstring(xml_string)




    # Iterate over <record> elements
    for record in root.findall('.//record'):

        # Check for physical format
        #if format_type in ["", "Physical", "physical"]:

        alma_title = record.findtext(".//datafield[@tag='245']/subfield[@code='a']").replace("/", "")

        # alma_title_b = record.findtext(".//datafield[@tag='245']/subfield[@code='b']")
        #
        # if alma_title_b is not None:
        #     alma_title_b = alma_title_b.replace("/", "")
        #     alma_title = alma_title + alma_title_b
        #

        alma_title = re.sub(r'[,:;."\'&]', '', alma_title)
        alma_title = re.sub(r'[-]', '', alma_title)

        alma_title = re.sub(r'\s*$', r'', alma_title)
        alma_title = re.sub(r'^(The|A|An)\s', r'', alma_title)
        # print('Alma title')
        #
        # print("-" + alma_title.lower() + "-")

        # print("open library title")
        open_library_title = re.sub(r'[,:;."\'&]', '', open_library_title)
        open_library_title = re.sub(r'[-]', '', open_library_title)

        open_library_title = re.sub(r'\s$', r'', open_library_title)

        open_library_title = re.sub(r'^(The|A|An)\s', r'', open_library_title)
        # print("Open library title")
        # print("-" + open_library_title.lower() + "-")



        author = record.findtext(".//datafield[@tag='100']/subfield[@code='a']") or \
                 record.findtext(".//datafield[@tag='110']/subfield[@code='a']") or \
                 record.findtext(".//datafield[@tag='700']/subfield[@code='a']") or \
                 record.findtext(".//datafield[@tag='710']/subfield[@code='a']")
        author = re.sub(r'(.+),$', r'\1', author)
        # print("Alma author")
        # print("-" + author + "-")
        # print("open library author")
        open_library_author = re.sub(r'.*?([^\s]+)$', r'\1', open_library_author)
        # print("-" + open_library_author + "-")


        # Convert the XML element back to a string
        pretty_xml = ET.tostring(record, encoding="unicode")

        # Print the pretty XML
        # print(pretty_xml)


        if (alma_title.lower() == open_library_title.lower() and open_library_title != "" and open_library_author in author and open_library_author != ""):

            # print("record type")
            # print(record.find(".//datafield[@tag='AVA']"))
            # print(record.find(".//datafield[@tag='AVE']"))
            # print("\n\n\n")

            ava = record.find(".//datafield[@tag='AVA']")
            if ava is not None:
                # print("AVA\n\n\n")
                phys_mms_id = ava.find("subfield[@code='0']").text

                # print("phsysical_mms_id")
                # print(phys_mms_id)
                # Simulate an API call to get item details
                item_query = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/{phys_mms_id}/holdings/ALL/items?apikey={api_bib}&format=json"
                response_items = requests.get(item_query)

                if response_items.status_code == 200:
                    items = response_items.json()

                    # print(json.dumps(items))
                    for item in items.get('item', []):
                        library = item.get('holding_data', {}).get('temp_library', {}).get('desc', '') if item.get('holding_data', {}).get('in_temp_location') else item.get('item_data', {}).get('library', {}).get('desc', '')
                        location = item.get('holding_data', {}).get('temp_location', {}).get('desc', '') if item.get('holding_data', {}).get('in_temp_location') else item.get('item_data', {}).get('location', {}).get('desc', '')

                        if item['item_data']['barcode'] + json_course.get('course', [{}])[0].get('code', '') + "-" + json_course.get('course', [{}])[0].get('section', '') in barcode_array:
                            continue
                        else:
                            barcode_array.append(item['item_data']['barcode'] + json_course.get('course', [{}])[0].get('code', '') + "-" + json_course.get('course', [{}])[0].get('section', ''))

                        author = record.findtext(".//datafield[@tag='100']/subfield[@code='a']") or \
                                 record.findtext(".//datafield[@tag='700']/subfield[@code='a']") or \
                                 record.findtext(".//datafield[@tag='710']/subfield[@code='a']")

                        # print("barcode array")
                        # print(barcode_array)
                        # print("\n\n\n\n\n")
                        results.append({
                            'Title': record.findtext(".//datafield[@tag='245']/subfield[@code='a']").replace("/", ""), #str(record.findtext(".//datafield[@tag='245']/subfield[@code='a']")).replace("None", "") +
                                    # str(record.findtext(".//datafield[@tag='245']/subfield[@code='b']")).replace("None", ""),
                            'Author': author,
                            'Publisher': record.findtext(".//datafield[@tag='264']/subfield[@code='b']"),
                            'Year': record.findtext(".//datafield[@tag='264']/subfield[@code='c']"),
                            'MMS ID': phys_mms_id,
                            'ISBN': record.findtext(".//datafield[@tag='020']/subfield[@code='a']"),
                            'Course Code': json_course.get('course', [{}])[0].get('code', ''),
                            'Course Section': json_course.get('course', [{}])[0].get('section', ''),
                            'Library': library,
                            'Location': location,
                            'Call Number': item.get('holding_data', {}).get('permanent_call_number', ''),
                            'Barcode': item.get('item_data', {}).get('barcode', ''),
                            'Description': json.dumps(item.get('item_data', {}).get('description', ''), ensure_ascii=False),
                            'Returned Format': 'Physical'
                        })

                        new_dict = {}
                        for key, value in row.items():


                            new_dict[key +  " - Input"] = value
                        old_dict = results[-1]



                        results[-1] = {**new_dict, **old_dict}

                        print(results[-1]['Course Code'] + "-" + results[-1]['Course Section'] + "-" + results[-1]['Title'] + "-" + results[-1]['MMS ID'])        # Check for electronic format

                else:
                    print("item retrieval problem")
                    results.append([{
                        'Title': f'Item API call failure for {row.get("TITLE", "")}, {item_query}',
                        'Author': f'Item API call failure for results for {row.get("AUTHOR", "")}, {item_query}',
                        'Year': row.get('Year', ''),
                        'Course Code': json_course.get('course', [{}])[0].get('code', ''),
                        'Returned Format': 'N/A'
                    }])

        # if format_type in ["", "Electronic", "electronic"]:
            ave = record.find(".//datafield[@tag='AVE']")
            if ave is not None:
                # print("AVE\n\n\n")
                e_mms_id = ave.find("subfield[@code='0']").text

                if f"{e_mms_id}Electronic" in electronic_record_array:
                    continue
                else:
                    electronic_record_array.append(f"{e_mms_id}Electronic")



                results.append({
                    'Title': alma_title,
                    'Author': author,
                    'Publisher': record.findtext(".//datafield[@tag='264']/subfield[@code='b']"),
                    'Year': record.findtext(".//datafield[@tag='264']/subfield[@code='c']"),
                    'MMS ID': e_mms_id,
                    'ISBN': record.findtext(".//datafield[@tag='020']/subfield[@code='a']"),
                    'Course Code': json_course.get('course', [{}])[0].get('code', ''),
                    'Course Section': json_course.get('course', [{}])[0].get('section', ''),
                    'Returned Format': 'Electronic'
                })


                new_dict = {}
                for key, value in row.items():


                    new_dict[key +  " - Input"] = value
                old_dict = results[-1]



                results[-1] = {**new_dict, **old_dict}

                print(results[-1]['Course Code'] + "-" + results[-1]['Course Section'] + "-" + results[-1]['Title'] + "-" + results[-1]['MMS ID'])
        else:

            results.append({
                'Title': f'No results for {row.get("TITLE", "")}',
                'Author': f'No results for {row.get("AUTHOR", "")}',
                'Year': row.get('Year', ''),
                'Course Code': json_course.get('course', [{}])[0].get('code', ''),
                'Returned Format': 'N/A'
            })

            new_dict = {}
            for key, value in row.items():


                new_dict[key +  " - Input"] = value
            old_dict = results[-1]



            results[-1] = {**new_dict, **old_dict}

    #except:

        #
        # results = [{'Title': 'NO RESULTS',
        # 'Author': 'NO RESULTS',
        # 'Contributor': 'NO RESULTS',
        # 'Publisher': r'NO RESULTS',
        # 'Date': 'NO RESULTS',
        # 'MMS ID': 'NO RESULTS',
        # 'ISBN': 'NO RESULTS',
        # 'Version': 'NO RESULTS',
        # 'Course Code': 'NO RESULTS',
        # 'Course Section': 'NO RESULTS',
        # 'Returned Format': 'NO RESULTS',
        # 'Library': 'NO RESULTS',
        # 'Location': 'NO RESULTS',
        # 'Call Number': 'NO RESULTS',
        # 'Barcode': 'NO RESULTS',
        # 'Description': 'NO RESULTS',
        # 'Citation Type': '',
        # 'section_info': 'NO RESULTS',
        # 'Item Policy': 'NO RESULTS'
        # }]
        #
        # for key, value in row.items():
        #     if key not in results[-1]:
        #         results[-1][key + " - Input"] = value
        # print(results[-1]['Course Code'] + "-" + results[-1]['Course Section'] + "-" + results[-1]['Title'] + "-" + results[-1]['MMS ID'])



    return results

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
def get_open_library_results(url):

    return requests.get(url)
def search_oclc(row):
    #title = row.get('Title', '')
    # author_last = row.get('Author', '')
    #format_type = row.get('Format', '')
    #course_number = row.get('Course Number', '')
    #course_semester = row.get('Term', '')
    oclc = str((row.get('OCLC Control Number (035a)', '')))

    # exploded_df = row.copy()
    # exploded_df['ISBN'] = exploded_df['ISBN'].split(';')
    # #
    # exploded_oclc_df = exploded_oclc_df.explode('OCLC').reset_index(=True)
    url_prefix = "http://www.worldcat.org/webservices/catalog/content"

    access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)

    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    



    #oclc_results = requests.get("http://api.crossref.org/works?query.bibliographic=" + str(isbn) + "&rows=1")
    #openlibrary_url = "https://openlibrary.org/api/books?bibkeys=ISBN:" + str(isbn) + "&jscmd=details&format=json"

    oclc_results = ""
    max_retry_range = 10  # Reduced for clarity, increase as needed

    for retry in range(max_retry_range):
        try:
            # Assuming get_open_library_results is a function that makes an API call
            oclc_results = get_open_library_results(openlibrary_url)
            break  # Success! Exit the loop
        except requests.exceptions.Timeout:  # Catch the correct timeout exception
            print(f"Attempt {retry + 1} failed. Retrying after waiting.")
            time.sleep((retry * 2) + 30)  # Exponential backoff
        except Exception as e:
            print(f"An error occurred: {e}")
            break  # Exit on other exceptions


    if openlibrary_results != "" and openlibrary_results.status_code == 200 and "ISBN:" + str(isbn) in openlibrary_results.json():
        openlibrary_results = openlibrary_results.json()
        #print(json.dumps(oclc_results))

        #print(oclc_results['message']['items'][0]['title'][0])
        #for item in oclc_results['message']['items']:




        #if item['type'] == 'monograph' or item['type'] == 'edited-book':
        #title = item['volumeInfo']['title']
        openlibrary_results_key = openlibrary_results["ISBN:" + str(isbn)] #.details.title
        # try:
        #
        #     title = openlibrary_results_key['details']['full_title']
        # except:
        title = openlibrary_results_key['details']['title']
        #
        print(openlibrary_results_key)
        return openlibrary_results_key
        #print(title)
        #print(isbn)

        # results = [{
        #         'Title': title,
        #         'Author': f'Not enough metadata  for {author_last}',
        #         #'Publisher': f'Not enough metadata  for {Publisher}',
        #         'Course Code': f'Not enough metadata for {course_number}',
        #         'Returned Format': 'N/A'
        #     }]

        

    #     publisher = ""
    #     author_first = ""
    #     author = ""
    #     try:
    #         author = openlibrary_results_key['details']['authors'][0]['name']

    #         author_first = re.sub(r'(^[^\s]+)\s.+', r'\1', author)
    #         author_last = re.sub(r'.+\s([^\s]+)$', r'\1', author)


    #     except:
    #         try:
    #             author_last = author

    #             #publisher = openlibrary_results_key['details']['publishers'][0]

    #         except:

    #             if author_last != "":
    #                 author_last = author

    #             else:
    #                 author_last = ""



    #     #print("\n" + publisher)
    #     #sys.exit()


    #     #if 'F' in course_semester:
    #     #    course_semester = course_semester.replace('F', 'Fa')
    #     #
    #     #elif 'W' in course_semester:
    #     #    course_semester = course_semester.replace('W', 'Sp')
    #     #instructor = row.get('Instructor Last Name', '')

    #     query = "https://tufts.alma.exlibrisgroup.com/view/sru/01TUN_INST?version=1.2&operation=searchRetrieve&recordSchema=marcxml&alma.mms_material_type=BK"
    #     open_library_title = title

    #     # note that per SRU instructions, the double equals for title is being used for titles that may contain a dash or an underline
    #     if open_library_title and open_library_title != "":

    #         open_library_title = re.sub(r'[,:;."\'&]', '', open_library_title)
    #         open_library_title = re.sub(r'[-]', '', open_library_title)

    #         query += f"&query=alma.title=%22{requests.utils.quote(title)}%22"
    #     else:

    #         results =  [{
    #             'Title': f'No results for {title}',
    #             'Author': f'No results for {author_last}',
    #             'Publisher': row.get('Publisher', ''),
    #             'Year': row.get('Year', ''),
    #             'Course Code': course_number,
    #             'Returned Format': 'N/A'
    #         }]

    #         new_dict = {}
    #         for key, value in row.items():


    #             new_dict[key +  " - Input"] = value
    #         old_dict = results[-1]



    #         results[-1] = {**new_dict, **old_dict}

    #         return results
    #     if author_last != "":

    #         #author_first = row.get('Author First', '')
    #         if author_first != author and author_first != "":
    #             query += f"%20AND%20alma.creator=%22*{requests.utils.quote(author_last + ',' + author_first)}*%22"
    #         else:
    #             query += f"%20AND%20alma.creator=%22*{requests.utils.quote(author_last)}*%22"
    #     # elif publisher != "":
    #     #     query += f"%20AND%20alma.publisher=%22*{requests.utils.quote(publisher)}*%22"
    #     else:

    #         results = [{
    #             'Title': f'Not enough metadata for results for {title}',
    #             'Author': f'Not enough metadata  for {author_last}',
    #             #'Publisher': f'Not enough metadata  for {Publisher}',
    #             'Course Code': f'Not enough metadata for {course_number}',
    #             'Returned Format': 'N/A'
    #         }]

    #         new_dict = {}
    #         for key, value in row.items():


    #             new_dict[key +  " - Input"] = value
    #         old_dict = results[-1]



    #         results[-1] = {**new_dict, **old_dict}

    #         return results


    #     response = requests.get(query)



    #     if response.status_code == 200:


    #         # Removing namespaces


    #         xml_string = re.sub(r'xmlns[^=]*="[^"]*"', '', response.text)
    #         xml_string = re.sub(r'[a-zA-Z]+:([a-zA-Z]+[=>])', r'\1', xml_string)

    #         root = ET.fromstring(xml_string)

    #         if root.findtext('.//numberOfRecords') != "0":
    #         # Simulate the course API response


    #             course_url = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey={api_key_courses}&q=name~{course_semester}-{course_number}%20AND%20instructors~{instructor}&format=json"

    #             course_url_without_instructor = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/courses?apikey={api_key_courses}&q=name~{course_semester}-{course_number}&format=json"

    #             response_course = ""
    #             if instructor != "":

    #                 response_course = requests.get(course_url)
    #             else:
    #                 response_course = requests.get(course_url_without_instructor)



    #             json_course = response_course.json() if response_course.status_code == 200 else {}


    #             # Parse the XML response
    #             return parse_xml_response(xml_string, json_course, format_type, row, api_bib, open_library_title, author)
    #         else:
    #             print("No Alma results")
    #             results = [{
    #                 'Title': f'No Alma results for {title}',
    #                 'Author': f'No Alma results for {author_last}',
    #                 'Course Code': f'No Alma results for {course_number}',
    #                 'Returned Format': 'N/A'
    #             }]

    #             new_dict = {}
    #             for key, value in row.items():


    #                 new_dict[key +  " - Input"] = value
    #             old_dict = results[-1]



    #             results[-1] = {**new_dict, **old_dict}

    #             return results
    #     else:
    #         print(response.status_code)
    #         print(response.text)
    #         results = [{
    #             'Title': f'No Alma results for {title}',
    #             'Author': f'No Alma results for {author_last}',
    #             'Publisher': f'No Alma results for {Publisher}',
    #             'Course Code': f'No Alma results for {course_number}',
    #             'Returned Format': 'N/A'
    #         }]
    #         new_dict = {}
    #         for key, value in row.items():


    #             new_dict[key +  " - Input"] = value
    #         old_dict = results[-1]



    #         results[-1] = {**new_dict, **old_dict}
    #         return results
    # # else:
    # #     print(json.dumps(oclc_results))
    else:

        print(openlibrary_results.status_code)
        print(openlibrary_results.text)
        results = [{
            'Title': f'No Open Library result for {row.get("TITLE", "")}',
            'Author': f'No Open Library  result for  {row.get("AUTHOR", "")}',
            'Contributor': f'No Open Library  result for {row.get("Contributor", "")}',
            'Year': f'No Open Library  result for {row.get("Year", "")}',
            'Course Code': f'No Open Library  result for {row.get("Course Number", "")}'
        }]

        new_dict = {}
        for key, value in row.items():


            new_dict[key +  " - Input"] = value
        old_dict = results[-1]



        results[-1] = {**new_dict, **old_dict}
        return results
  
def process_spreadsheet(df):
    #df = read_excel_file(file_path)
    #df = df[(df['EAN-13'] != "") & (~df['EAN-13'].isna())]

    print(len(df))

    print("\n\n\n")
    results = []
    #df = df.rename(columns={'Course': 'Numeric'})
    #df = df.rename(columns={'Title': 'TITLE'})
    #df = df.rename(columns={'Author': 'AUTHOR'})
    #pd.set_option('display.max_columns', None)




    # #df['Course Number'] = df['Dept'].apply(lambda x: x) + "-" +  df['Numeric'].apply(lambda x: x) + "-" + df['Sec'].apply(lambda x: x)


    # exploded_df = df
    # exploded_df['ISBN'] = exploded_df['ISBN'].str.split(';')

    
    # exploded_df = exploded_df.explode('ISBN').reset_index()




    x = 0
    for _, row in df.iterrows():
        # if x == 25:
        #     break
        if x != 0 and x % 100 == 0:
            time.sleep(30)
        x += 1

        if x == 30:
            break
        result = search_oclc(row)
        if result:
            results.extend(result)
        else:
            results.append({
                'Title': f'No results for {row.get("TITLE", "")}',
                'Author': f'No results for {row.get("AUTHOR", "")}',
                'Contributor': f'No results format for {row.get("Contributor", "")}',
                'Year': f'No results for {row.get("Year", "")}',
                'Course Code': f'No results {row.get("Course Number", "")}',
                'IBSN': f'No results for {row.get("EAN-13", "")}'
            })

    print(results)
    return results




def generate_excel(results):
    data = []

    def process_result(result):
        # print("before")
        # print(result)

        try:
            results = {
                'Title': result.get('Title', ''),
                'Author': result.get('Author', ''),
                'Contributor': result.get('Contributor', ''),
                'Publisher': result.get('Publisher', ''),
                'Date': result.get('Year', ''),
                'MMS ID': result.get('MMS ID', ''),
                'ISBN': result.get('ISBN', ''),
                'Version': result.get('Version', ''),
                'Course Code': result.get('Course Code', ''),
                'Course Section': result.get('Course Section', ''),
                'Returned Format': result.get('Returned Format', ''),
                'Library': result.get('Library', ''),
                'Location': result.get('Location', ''),
                'Call Number': result.get('Call Number', ''),
                'Barcode': result.get('Barcode', ''),
                'Description': result.get('Description', '')
                #'Citation Type': '',
                #'section_info': result.get('section_info', ''),
                #'Item Policy': result.get('item_policy', result.get('Item Policy', ''))
            }
        except:
            results = {'Title': 'NO RESULTS',
            'Author': 'NO RESULTS',
            'Contributor': 'NO RESULTS',
            'Publisher': r'NO RESULTS',
            'Date': 'NO RESULTS',
            'MMS ID': 'NO RESULTS',
            'ISBN': 'NO RESULTS',
            'Version': 'NO RESULTS',
            'Course Code': 'NO RESULTS',
            'Course Section': 'NO RESULTS',
            'Returned Format': 'NO RESULTS',
            'Library': 'NO RESULTS',
            'Location': 'NO RESULTS',
            'Call Number': 'NO RESULTS',
            'Barcode': 'NO RESULTS',
            'Description': 'NO RESULTS',
            'Citation Type': '',
            'section_info': 'NO RESULTS',
            'Item Policy': 'NO RESULTS'
            }

        # print("after")
        # print(results)
        #print(row['course_code'] + "-" + row['Title'] + "-" + row['MMS ID'])
        # Add any extra fields not already included in the row
        # new_values = {k: v for k, v in results.items() if k not in result}
        # new_row = {**new_values, **result}
        # print("after")
        # print(results)
        # print("\n\n\n")
        #return new_row

    # Process each result
    if isinstance(results, list):
        print("list")

        for result in results:
            # print(result)
            # result = { k: ('' if v is None else v) for k, v in result.items() }
            # result = json.loads(result)
            data.append(result)
    else:
        # print("not list")
        # print(results)
        # result = { k: ('' if v is None else v) for k, v in result.items() }
        # result = json.loads(results)
        data.append(results)


    # result = { k: ('' if v is None else v) for k, v in some_dict.items() }
    # Convert to DataFrame
    # print(data)
    result_columns = []
    for k in data[0]:
        # print("data type of list dict 'row'")
        # print(type(data[0]))


        result_columns.append(k)

    # print(result_columns)

    df = pd.DataFrame(columns=result_columns)

    error_file_string = ""
    for list_row in data:
        # print(list_row)
        try:

            df = pd.concat([df, pd.DataFrame(list_row, index=[0])])
        except:
            error_file_string = str(list_row) + "\n"

    df = df.reset_index()
    # print(df)
    # sys.exit()
    # df = pd.DataFrame(data)

    # print(df)
    # Create an Excel workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = 'Results'

    # Write DataFrame to Excel sheet
    for r_idx, row in df.iterrows():
        for c_idx, value in enumerate(row):
            cell = ws.cell(row=r_idx + 2, column=c_idx + 1, value=value)

            # Handle Barcode formatting
            if df.columns[c_idx] == 'Barcode':
                cell.number_format = numbers.FORMAT_TEXT

    # Set the column headers
    for idx, col in enumerate(df.columns, 1):
        ws.cell(row=1, column=idx, value=col)

    # Save the workbook to a file
    oDir = "Barnes and Noble Parsed"
    
    if not os.path.isdir(oDir) or not os.path.exists(oDir):
        os.makedirs(oDir)
    
    error_file = open("Open Library Error File.txt", "w+")

    error_file.write(error_file_string)

    error_file.close()
    wb.save(oDir + '/Barnes and Noble Data with Matching Inventory and Course Codes -v3.xlsx')

def to_list(v):
    if isinstance(v, list):
        return v
    if pd.isna(v):
        return []
    # if semicolon-separated string, split; otherwise wrap as single-item list
    s = str(v)
    return [x.strip() for x in s.split(';')] if ';' in s else [s]

def main():

    import time

    start_time = time.perf_counter()
    # Sample results for demonstration
    # results = [
    #     {
    #         'Title': 'Sample Title',
    #         'Author': 'Sample Author',
    #         'Contributor': 'Sample Contributor',
    #         'Publisher': 'Sample Publisher',
    #         'Year': '2024',
    #         'MMS ID': '123456',
    #         'ISBN': '978-3-16-148410-0',
    #         'Version': '1.0',
    #         'course_code': 'CS101',
    #         'course_section': '001',
    #         'Returned Format': 'Physical',
    #         'Library': 'Main Library',
    #         'Location': 'Floor 2',
    #         'Call Number': 'QA76.73.P98 G8 2024',
    #         'Barcode': json.dumps('0123456789'),
    #         'Description': json.dumps('Sample description.'),
    #         'section_info': 'Info Section',
    #         'item_policy': 'Standard'
    #     }
    # ]

    output_dir = os.path.join(os.getcwd(), "Output")
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    input_dir = os.path.join(os.getcwd(), "input")
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
    semester_counts_file_path = filedialog.askopenfilename(title="Upload Excel file of semester course counts")

    academic_deparment_filepath = filedialog.askopenfilename(title="Upload Excel file of academic departments")
    
    sushi_counts_file_path = filedialog.askopenfilename(title="Upload Excel file of SUSHI ebook counts")# = semester_counts_file_path.replace("\\", "/")
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



    # semester_df = semester_df.explode('ISBN (Normalized)').reset_index()
    # semester_df["ISBN (Normalized)"] = semester_df["ISBN (Normalized)"].apply(lambda x: x.strip() if isinstance(x, str) else x)
    # academic_df['ISBN (Normalized)'] = academic_df['ISBN (Normalized)'].str.split(';')
    # academic_df = academic_df.explode('ISBN (Normalized)').reset_index()
    # academic_df["ISBN (Normalized)"] = academic_df["ISBN (Normalized)"].apply(lambda x: x.strip() if isinstance(x, str) else x)
    merged_df = pd.merge(semester_df, academic_df, how='right', left_on=['MMS Id', 'Associated Course Code'], right_on=['MMS Id', 'Course Code'])
    #merged_df = pd.merge(semester_df, academic_df, how='left', left_on=['MMS Id', 'Course Code'], right_on=['MMS Id', 'Course Code'])
    merged_df = merged_df.reset_index(drop=True)  
    merged_df.to_csv("Output/Merged Semester Loans and Course Subjects Unrolled.csv", index=False)
    #sys.exit()
    # Example: merge if left['A'] OR left['B'] matches right['X']

    merged_df['ISBN (Normalized)'] = merged_df['ISBN (Normalized)_x'].combine_first(merged_df['ISBN (Normalized)_y'])   
# then group by the single 'ISBN (Normalized)'

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
    #file_path = reserves_files[0]
    #file_path = askopenfilename(title="Upload Excel file from bookstore")
    #api_key_courses = secrets_local.prod_courses_api_key

    #api_bib = secrets_local.api_bib

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

    grouped_oclc_df.to_csv("Output/Grouped by OCLC Control Number.csv", index=False)
    #sys.exit()

    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    print(f"Elapsed Time: {elapsed_time} seconds")
    sys.exit()
    results = process_spreadsheet(merged_with_sushi_df)


    # Generate Excel file
    generate_excel(results)

if __name__ == "__main__":
    main()
