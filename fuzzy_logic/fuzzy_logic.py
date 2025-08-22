"""
Author: Yuewen Li
Date: 2025/07/25
Description: this script use fuzzy logic to help users find the correct researcher name when they accidentally make
typos or miss some words. It can also help user to find the correct researcher name when they only enter the first name or last name.
If multiple matches are found, the script will ask the user to provide more information about the researcher to get the correct one.


Requirements:
To install the required libraries, use the following commands:
    pip install rapidfuzz
    pip install pandas
    pip install openpyxl


"""


import pandas as pd
from rapidfuzz import fuzz, process


def fuzzy_match_multi(user_input, candidates, min_similarity=75, strict_contains=False):
    algorithms = [
        fuzz.token_set_ratio,
        fuzz.token_sort_ratio,
        fuzz.partial_ratio
    ]
    matches = []

    for algorithm in algorithms:
        result = process.extract(user_input, candidates, scorer=algorithm, score_cutoff=min_similarity, limit=None)
        matches.extend(result)

    final_matches = {}
    for candidate, score, idx in matches:
        # 如果 strict_contains 为 True，则只保留包含user input的结果
        # when strict_contains is false then not strict_contains is true
        # so that all fuzzy matches are included.
        # otherwise it will only have the results which contain the user input.
        if not strict_contains or user_input in candidate.lower():
            if candidate not in final_matches or score > final_matches[candidate][0]:
                final_matches[candidate] = (score, idx + 1)

    # sorted the result scores from high to low
    sorted_matches = sorted(final_matches.items(), key=lambda x: x[1][0], reverse=True)

    # return only index and name
    return [(candidate, idx)  for candidate, (score, idx) in sorted_matches]


def refine_matches(user_input, previous_matched,records, min_similarity=75):
    for candidate, idx in previous_matched:
        record = records.iloc[idx - 1]
        # If the email entered by the user exactly matches
        if user_input == record['Email_Address']:
            return [(record['Employee_Full_Name'], idx)]
        full_names = (records["First_Name"].astype(str).str.strip() + " " + records["Last_Name"].astype(
            str).str.strip()).str.lower()
        if any(user_input == name for name in full_names):
            return [(record['Employee_Full_Name'], idx)]

    subset_candidates = []
    idx_map = {}
    for candidate, idx in previous_matched:
        record = records.iloc[idx - 1]
        combined = f"{record['Employee_Full_Name']} {record['First_Name']} {record['Last_Name']} {record['Email_Address']} {record['Job_Name']} {record['Department']} {record['School']}".lower()
        subset_candidates.append(combined)
        idx_map[combined] = idx
    refine_match = fuzzy_match_multi(user_input, subset_candidates, min_similarity=73)
    return [(match, idx_map[match]) for match, _ in refine_match]


# read execel files and combine them into a dictionary list.
# each record contains name + email + job + department + school
records=pd.read_excel("reserchaer_names_details.xlsx",sheet_name='Sheet1')


# user input
user_input = input("Please enter a name or email: ").strip().lower()
#if the user input is email
if "@" in user_input:
    strict = any(user_input in str(email).lower() for email in records["Email_Address"])
    email_candidate = records["Email_Address"].astype(str).str.lower().tolist()
    matched = fuzzy_match_multi(user_input, email_candidate, strict_contains=strict)
    if matched:
        print(f"Matched email: {matched}")
        while len(matched) > 1:
            user_input = input("Multiple matches found. Please enter more information(department, name, job or school): ").strip().lower()
            matched = refine_matches(user_input,matched, records)
            print(f"Matched name: {matched}")
    else:
        print("No good match found for the email.")
else:
    # if the user input is name
    full_names = (records["First_Name"].astype(str).str.strip() + " " + records["Last_Name"].astype(
        str).str.strip()).str.lower()
    strict = any(user_input in name for name in full_names)
    name_candidate = full_names.tolist()

    matched = fuzzy_match_multi(user_input, name_candidate, strict_contains=strict)
    if matched:
        print(f"Matched name: {matched}")
        while len(matched) > 1:
            user_input = input("Multiple matches found. Please enter more information(department, email, job or school): ").strip().lower()
            matched = refine_matches(user_input, matched, records)
            print(f"Matched name: {matched}")
    else:
        print("No good match found for the name.")









