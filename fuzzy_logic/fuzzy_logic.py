# install rapidfuzz
#pip install rapidfuzz

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
        # 如果 strict_contains 为 True，则只保留包含关键词的结果
        if not strict_contains or user_input in candidate.lower():
            if candidate not in final_matches or score > final_matches[candidate][0]:
                final_matches[candidate] = (score, idx + 1)

    sorted_matches = sorted(final_matches.items(), key=lambda x: x[1][0], reverse=True)
    return [(candidate, idx) for candidate, (score, idx) in sorted_matches]



    # sorted_indices = sorted(final_matches.values())
    # return [i + 1 for i in sorted_indices]

with open("names.txt", "r", encoding="utf-8") as file:
    names_list = [line.strip() for line in file if line.strip()]

names_list_lower = [name.lower() for name in names_list]

with open("emails.txt", "r", encoding="utf-8") as file:
    emails_list = [line.strip() for line in file if line.strip()]

emails_list_lower = [email.lower() for email in emails_list]

user_input = input("Please enter a name or email: ").strip().lower()



if "@" in user_input:
    strict = any(user_input in email for email in emails_list_lower)
    matched = fuzzy_match_multi(user_input, emails_list, strict_contains=strict)
    if matched:
        print(f"Matched email: {matched}")
    else:
        print("No good match found for the email.")
else:
    strict = any(user_input in name for name in names_list_lower)
    matched = fuzzy_match_multi(user_input, names_list_lower, strict_contains=strict)
    if matched:
        print(f"Matched name: {matched}")
    else:
        print("No good match found for the name.")






