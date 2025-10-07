// ...existing code...
def search_oclc(oclc):
    """
    Query OCLC for a single OCLC number and return a dict with keys:
    'title', 'dewey', 'pub_date', 'oclc_number' (strings, empty if not found).
    Uses global access_token. Handles basic retries/backoff.
    """
    global access_token

    oclc_str = str(oclc).strip()
    if not oclc_str or oclc_str.lower().startswith("nan"):
        return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}

    url_prefix = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs?q=no:"
    url = f"{url_prefix}{oclc_str}"

    headers = {
        "accept": "application/json",
        "Authorization": f"Bearer {access_token}"
    }

    max_retries = 5
    backoff = 2
    for attempt in range(1, max_retries + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=30)
            if resp.status_code == 200:
                j = resp.json()
                # Attempt a few common extraction paths; fall back to empty strings
                title = ""
                dewey = ""
                pub_date = ""
                oclc_number = oclc_str

                # Many OCLC APIs return entries or records; try common keys.
                # entry list
                entries = j.get("entries") or j.get("entry") or j.get("records") or j.get("items") or []
                first = None
                if isinstance(entries, list) and len(entries) > 0:
                    first = entries[0]
                elif isinstance(entries, dict):
                    first = entries

                # Try a few likely places for title, dewey, date
                if first:
                    # title
                    title = first.get("title") or first.get("titleDisplay") or title
                    if isinstance(title, list):
                        title = title[0] if title else ""
                    # dewey
                    # sometimes classification->classificationIdentifier or classification->notation
                    if "classification" in first:
                        cls = first.get("classification")
                        if isinstance(cls, list) and cls:
                            c0 = cls[0]
                        else:
                            c0 = cls
                        if isinstance(c0, dict):
                            dewey = c0.get("notation") or c0.get("classificationIdentifier") or dewey
                    # pub date
                    pub_date = first.get("date") or first.get("published") or first.get("pubDate") or pub_date

                # fallback lookups across root object
                if not title:
                    title = j.get("title") or ""
                if not dewey:
                    # try scanning for 'dewey' or 'lc' strings
                    for k in ("dewey", "classification", "classifications"):
                        if k in j:
                            dv = j.get(k)
                            if isinstance(dv, str):
                                dewey = dv
                                break

                # normalize to strings
                title = "" if title is None else str(title)
                dewey = "" if dewey is None else str(dewey)
                pub_date = "" if pub_date is None else str(pub_date)

                return {"title": title, "dewey": dewey, "pub_date": pub_date, "oclc_number": oclc_number}
            elif resp.status_code in (429, 503):
                # rate limited or service unavailable -> backoff and retry
                time.sleep(backoff)
                backoff *= 2
                continue
            else:
                # non-200 non-retryable -> return empty values
                return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}
        except requests.exceptions.Timeout:
            time.sleep(backoff)
            backoff *= 2
            continue
        except Exception:
            # unexpected parsing/network error -> return empties
            return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}

    # if all retries exhausted
    return {"title": "", "dewey": "", "pub_date": "", "oclc_number": oclc_str}
// ...existing code...
def process_spreadsheet(oclc_list, output_csv_path, chunk_sleep=0.25, write_header_once=True):
    """
    Iterate over oclc_list and for each OCLC:
      - query OCLC (search_oclc)
      - find corresponding rows in grouped_oclc_df
      - set OCLC columns on those rows
      - append the updated rows to output_csv_path incrementally

    This prevents holding all merged rows in memory and persists progress as we go.
    """

    global grouped_oclc_df, access_token

    # ensure control numbers are normalized for matching
    grouped_oclc_df["__OCLC_norm__"] = grouped_oclc_df["OCLC Control Number (035a)"].astype(str).str.strip()

    # remove existing output file so we start fresh (only if we want to)
    if os.path.exists(output_csv_path):
        os.remove(output_csv_path)

    header_written = False

    total = len(oclc_list)
    for idx, raw_oclc in enumerate(oclc_list, start=1):
        oclc_str = str(raw_oclc).strip()
        if not oclc_str or oclc_str.lower().startswith("nan"):
            # still append matching rows (if any) with blank OCLC fields
            matched = grouped_oclc_df[grouped_oclc_df["__OCLC_norm__"] == oclc_str]
            if not matched.empty:
                matched = matched.copy()
                matched["OCLC TitleAuthor"] = ""
                matched["OCLC Dewey"] = ""
                matched["OCLC Publication Date"] = ""
                matched["OCLC Number"] = oclc_str
                # write/append
                matched.to_csv(output_csv_path, mode="a", index=False, header=not header_written)
                header_written = True
            continue

        # get OCLC info
        info = search_oclc(oclc_str)

        # find matching rows in original DF
        matched = grouped_oclc_df[grouped_oclc_df["__OCLC_norm__"] == oclc_str]
        if matched.empty:
            # nothing to append for this OCLC; continue
            print(f"[{idx}/{total}] OCLC {oclc_str}: no matching rows in input.")
        else:
            matched = matched.copy()
            matched["OCLC TitleAuthor"] = info.get("title", "")
            matched["OCLC Dewey"] = info.get("dewey", "")
            matched["OCLC Publication Date"] = info.get("pub_date", "")
            matched["OCLC Number"] = info.get("oclc_number", oclc_str)

            # append to CSV (write header only once)
            matched.to_csv(output_csv_path, mode="a", index=False, header=not header_written)
            header_written = True
            print(f"[{idx}/{total}] Appended {len(matched)} row(s) for OCLC {oclc_str}")

        # light throttle between requests; every N requests you may want longer sleep
        time.sleep(chunk_sleep)

    # cleanup helper column
    if "__OCLC_norm__" in grouped_oclc_df.columns:
        grouped_oclc_df.drop(columns="__OCLC_norm__", inplace=True)

// ...existing code...
def main():
    global grouped_oclc_df

    start_time = time.perf_counter()

    # ensure OCLC result columns exist
    grouped_oclc_df['OCLC TitleAuthor'] = ""
    grouped_oclc_df['OCLC Dewey'] = ""
    grouped_oclc_df['OCLC Publication Date'] = ""
    grouped_oclc_df['OCLC Number'] = ""

    # normalize control number column
    grouped_oclc_df['OCLC Control Number (035a)'] = grouped_oclc_df['OCLC Control Number (035a)'].astype(str).str.strip()

    number_of_mms_ids = len(grouped_oclc_df['MMS Id'].unique())
    number_of_oclcs = len(grouped_oclc_df['OCLC Control Number (035a)'].unique())

    unique_oclcs = grouped_oclc_df['OCLC Control Number (035a)'].unique()
    print(f"Number of unique MMS IDs: {number_of_mms_ids}")
    print(f"Number of unique OCLC Control Numbers: {number_of_oclcs}")

    # prepare output filename with timestamp
    from datetime import datetime
    now = datetime.now()
    timestamp_filename = now.strftime("%Y%m%d%H%M%S")
    output_filename = f"Output/Library Statistics with Merged OCLC Data {timestamp_filename}.csv"
    os.makedirs(os.path.dirname(output_filename), exist_ok=True)

    oclc_answer = input("Do you already have OCLC data (yes/no)?: ").strip().lower()
    oclc_bool = (oclc_answer == "yes")

    if oclc_bool:
        print("Using existing OCLC data in grouped_oclc_df; writing out immediately.")
        # write the existing dataframe directly (no incremental enrich)
        grouped_oclc_df.to_csv(output_filename, index=False)
    else:
        # ensure we have an access token before starting
        access_token = get_oclc_token(secrets_local.client_id, secrets_local.client_secret)
        # process and write incrementally
        process_spreadsheet(unique_oclcs, output_filename, chunk_sleep=0.25)

    end_time = time.perf_counter()
    elapsed_time = end_time - start_time
    print(f"Elapsed Time: {elapsed_time:.1f} seconds")
    print(f"Final output written to: {output_filename}")
// ...existing code...