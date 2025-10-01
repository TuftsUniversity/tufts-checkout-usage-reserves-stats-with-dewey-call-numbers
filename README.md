# Process Tufts loan, reserves reading lists, and "SUSHI" publisher electronic usage data with course subjects, and OCLC-provided Dewey decimal numbers to analyze subject area usage at Tufts


Author: Henry Steele, Library Technology Services, Tufts University


## Input data
- This process relies on three input files source from Alma Analytics.  This data has to be provided by library staff because non-library staff don't have access to it, but below I have provided the SQL queries that generate these reports so anyone with such access can easily create them.  I've also provided their paths in the Tufts Libraries Alma Analytics platform
    - 	/shared/Tufts University/Reports/Fulfillment/Circulation Statistics/Semester Reserve Stats by Course
    -	/shared/Tufts University/Reports/Fulfillment/Circulation Statistics/Academic Department for Reserves and MMS ID
    -	/shared/Tufts University/Reports/Fulfillment/Circulation Statistics/	SUSHI Usage Stats by Date for Ebooks - Fiscal Years During Which We Had Canvas Reading List Tool
- the report generator SQL also lists the field reqiurements
    - "Semester Reserve Stats by Course"
        `SELECT 
        "Item Location at time of loan"."Library Name" saw_0,
        "Item Location at time of loan"."Location Name" saw_1,
        "Physical Item Details"."Item Policy" saw_2,
        "Loan Date"."Loan Fiscal Year" saw_3,
        "Bibliographic Details"."MMS Id" saw_4,
        "Bibliographic Details"."OCLC Control Number (035a)" saw_5,
        "Bibliographic Details"."ISBN (Normalized)" saw_6,
        "Loan Details"."Associated Course Code" saw_7,
        'FY-'||Evaluate('Regexp_substr(%1,''^(\d)(\d\d)\d-'',1,1,''i'',1)',"Loan Details"." Associated Course Code")||'0'|| Evaluate('Regexp_substr(%1,''^(\d)(\d\d)\d-'',1,1,''i'',2)',"Loan Details"." Associated Course Code") saw_8,
        "Loan"."Loans (Not In House)" saw_9
        FROM "Fulfillment"
        WHERE 
        ("Loan Details"."In House Loan Indicator" = 'N') AND ("Loan Date"."Loan Fiscal Year" IN ('FY-2018', 'FY-2019', 'FY-2020', 'FY-2021', 'FY-2022', 'FY-2023', 'FY-2024', 'FY-2025', 'FY-2026')) AND ("Loan Details"."Associated Course Code" IS NOT NULL)
        `
    - "Academic Department for Reserves and MMS ID"
        `SELECT
            0 s_0,
            "Fulfillment"."Bibliographic Details"."ISBN (Normalized)" s_1,
            "Fulfillment"."Bibliographic Details"."MMS Id" s_2,
            "Fulfillment"."Bibliographic Details"."OCLC Control Number (035a)" s_3,
            "Fulfillment"."Item Location at time of loan"."Library Name" s_4,
            "Fulfillment"."Item Location at time of loan"."Location Name" s_5,
            "Fulfillment"."Loan Date"."Loan Fiscal Year" s_6,
            "Fulfillment"."Loan Details"."Associated Course Code" s_7,
            "Fulfillment"."Physical Item Details"."Item Policy" s_8,
            'FY-'||Evaluate('Regexp_substr(%1,''^(\d)(\d\d)\d-'',1,1,''i'',1)',"Fulfillment"."Loan Details"." Associated Course Code")||'0'||Evaluate('Regexp_substr(%1,''^(\d)(\d\d)\d-'',1,1,''i'',2)',"Fulfillment"."Loan Details"." Associated Course Code") s_9,
            "Fulfillment"."Loan"."Loans (Not In House)" s_10,
            REPORT_SUM("Fulfillment"."Loan"."Loans (Not In House)" BY ) s_11
            FROM "Fulfillment"
            WHERE
            (("Loan Details"."In House Loan Indicator" = 'N') AND ("Loan Date"."Loan Fiscal Year" IN ('FY-2018', 'FY-2019', 'FY-2020', 'FY-2021', 'FY-2022', 'FY-2023', 'FY-2024', 'FY-2025', 'FY-2026')) AND ("Loan Details"."Associated Course Code" IS NOT NULL))
            ORDER BY 7 ASC NULLS FIRST, 5 ASC NULLS FIRST, 6 ASC NULLS FIRST, 9 ASC NULLS FIRST, 8 ASC NULLS FIRST, 3 ASC NULLS FIRST, 10 ASC NULLS FIRST, 4 ASC NULLS FIRST, 2 ASC NULLS FIRST
            FETCH FIRST 10000001 ROWS ONLY
            `
    - "	SUSHI Usage Stats by Date for Ebooks - Fiscal Years During Which We Had Canvas Reading List Tool"
        - `SELECT
            0 s_0,
            "Usage Data"."Platform"."Platform" s_1,
            "Usage Data"."Title Identifier"."Display Title" s_2,
            "Usage Data"."Title Identifier"."Normalized ISBN" s_3,
            "Usage Data"."Title Identifier"."Normalized Title" s_4,
            "Usage Data"."Usage Data Details - Release 5"."Section Type" s_5,
            "Usage Data"."Usage Date"."Usage Date Fiscal Year" s_6,
            "Usage Data"."Usage Data Details - Release 5"."Title Identifier Count" s_7,
            "Usage Data"."Usage Data Details - Release 5"."Usage Measures Total" s_8,
            MAX(RCOUNT(8)) s_9
            FROM "Usage Data"
            WHERE
            (("Usage Data Details - Release 5"."Material Type Indicator" IN ('DR - Unique Title Requests', 'PR - Unique Title Requests', 'PR_P1 - Unique Title Requests', 'TR - Unique Title Requests', 'TR_B1 - Unique Title Requests', 'TR_B3 - Unique Title Requests')) AND ("Title Identifier"."Normalized ISBN" IS NOT NULL) AND ("Usage Date"."Usage Date Fiscal Year" IN ('FY-2018', 'FY-2019', 'FY-2020', 'FY-2021', 'FY-2022', 'FY-2023', 'FY-2024', 'FY-2025')))
            ORDER BY 2 ASC NULLS FIRST, 5 ASC NULLS FIRST, 3 ASC NULLS FIRST, 4 ASC NULLS FIRST, 6 ASC NULLS FIRST, 7 ASC NULLS FIRST
            FETCH FIRST 10000001 ROWS ONLY
        `
## Processing and Output
    - the script takes these inputs and transforms them such that they can be merged with each other, producing an intermediate result where loan data for physical items is matched with all course citations (that also include electronic).  This is a RIGHT JOIN.   
    - Then the ISBNs for each of these is "exploded" (separated into separate rows), and the course codes and course information is rolled up by fiscal year, so that then this can be compared with publisher supplied electronic usage stats
    - Note that the SUSH stats only contain one ISBN, but our records have many, so the exploded data allows any of these to be tried for a match
    - finally each row in this resultant data sheet is fed into an OCLC API call to get call number 
	
## Install and Run

	- install Python
		- https://www.python.org/downloads/
		- the installation and configuration of Python such that Python is available as a system command in your environement variables is outside the scope of this document
	- install requirements
		- `python3 -m pip install -r requirements.txt`
	- run program and produce outputs.  This may take a few minutes
		- `python3 getOCLCDewey.py`
		- for the initial process of merging the Alma data, the script runtime is around 2 minutes on a machine with 16GB of RAM running on roughly:
			- 3,000 rows of semester data
			- 8,000 rows of course data
			- 132,000 rows of SUSHI data
		- the part that retrieves Dewey Decimal Numbers from OCLC can take an hour or more on data of this size.
		- if you have already retreived the OCLC data, and have made modifications to the script and want to process with this data, you can answer "yes" to the prompt

## Data Model

![Data Model](Library Usage, Reserves Subjects, and Classifications Stats Data Moel.drawio.svg)