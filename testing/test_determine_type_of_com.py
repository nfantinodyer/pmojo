import datetime
import sqlite3
import json
import requests
from bs4 import BeautifulSoup

### ========== BEGIN: Imports from your main script ==========

# If you put these in another file, just import them. Otherwise, copy them inline.

def determine_type_of_com(cdi, cdn):
    """
    Original logic:
      if cdi in [23, 130]:
        if cdn == 1: return "e"
        elif cdn == 2: return "t"
        return ""
      else:
        if cdn == 1:
          if cdi in {1, 8, 21, 22, 30, 33, 35}: return "l"
          elif cdi == 36: return "e"
        elif cdn == 2:
          if cdi in {1, 8, 30, 33}: return "e"
        elif cdn == 3 and cdi == 1:
          return "t"
      return ""
    """
    if cdi in [23, 130]:
        if cdn == 1:
            return "e"
        elif cdn == 2:
            return "t"
        return ""
    else:
        if cdn == 1:
            if cdi in {1, 8, 21, 22, 30, 33, 35}:
                return "l"
            elif cdi == 36:
                return "e"
        elif cdn == 2:
            if cdi in {1, 8, 30, 33}:
                return "e"
        elif cdn == 3 and cdi == 1:
            return "t"
    return ""

def login_to_practice_mojo(username, password):
    """
    Logs into PracticeMojo and returns an authenticated requests.Session.
    """
    session = requests.Session()
    login_page_url = "https://app.practicemojo.com/Pages/login"
    session.get(login_page_url)

    login_url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/userLogin"
    form_data = {
        "loginId": username,
        "password": password,
        "slug": "login"
    }
    headers = {
        "Referer": login_page_url,
        "Content-Type": "application/x-www-form-urlencoded"
    }
    resp = session.post(login_url, data=form_data, headers=headers)
    if resp.status_code != 200:
        raise Exception(f"Login failed with status code {resp.status_code}.")

    if "Invalid login" in resp.text or "Incorrect password" in resp.text:
        raise Exception("Login failed: invalid credentials or site error.")

    return session

def fetch_activity_detail(session, date_str, cdi, cdn):
    """
    Fetch the detail page for a given date/cdi/cdn, return a list of row dicts
    in the exact DOM order (top to bottom).
    """
    base_url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/"
    detail_url = f"{base_url}gotoActivityDetail?td={date_str}&cdi={cdi}&cdn={cdn}"
    resp = session.get(detail_url)
    if resp.status_code != 200:
        raise Exception(f"Failed to fetch detail page: status={resp.status_code}")

    soup = BeautifulSoup(resp.text, "html.parser")
    return parse_activity_detail(soup)

def parse_activity_detail(soup):
    """
    Returns list of dicts: {
      "campaign": ...,
      "method": ...,
      "patient_name": ...,
      "appointment": ...,
      "confirmations": ...,
      "status": ...
    }, preserving table row order.
    """
    results = []
    table = soup.select_one("div.activity_detail_table table")
    if not table:
        return results

    rows = table.find_all("tr", recursive=False) or table.find_all("tr")
    i = 0
    while i < len(rows):
        tds = rows[i].find_all("td", recursive=False)
        if len(tds) >= 3 and tds[0].has_attr("rowspan") and tds[1].has_attr("rowspan"):
            campaign_name = tds[0].get_text(strip=True)
            method_name   = tds[1].get_text(strip=True)
            sub_count     = int(tds[0]["rowspan"])
            i += 1
            block_end = i + (sub_count - 1)

            while i < block_end and i < len(rows):
                data_row = rows[i]
                data_tds = data_row.find_all("td", recursive=False)
                if len(data_tds) == 4:
                    patient_name  = data_tds[0].get_text().rstrip()
                    appointment   = data_tds[1].get_text(strip=True)
                    confirmations = data_tds[2].get_text(strip=True)
                    row_status    = data_tds[3].get_text(strip=True)

                    patient_name = patient_name.replace("↳", "").rstrip()
                    if "Family" not in patient_name:
                        results.append({
                            "campaign": campaign_name,
                            "method": method_name,
                            "patient_name": patient_name,
                            "appointment": appointment,
                            "confirmations": confirmations,
                            "status": row_status
                        })
                i += 1
        else:
            i += 1

    return results

### ========== END: Imports from your main script ==========

#
# PART 1: Basic Unit Tests (local-only)
#

def run_unit_tests():
    """
    Tests the function 'determine_type_of_com' with known (cdi, cdn) -> expected_result.
    These are 'pure' function tests, not calling PracticeMojo.
    """
    test_cases = [
        # cdi=23, cdn=1 => 'e'
        (23, 1, "e"),
        (23, 2, "t"),
        (23, 3, ""),  # anything else => ''
        # cdi=130, cdn=1 => 'e'
        (130, 1, "e"),
        (130, 2, "t"),
        (130, 99, ""),
        # cdi=1, cdn=1 => 'l'
        (1, 1, "l"),
        (1, 2, "e"),
        (1, 3, "t"),
        # cdi=8, cdn=1 => 'l', cdn=2 => 'e'
        (8, 1, "l"),
        (8, 2, "e"),
        # cdi=36, cdn=1 => 'e'
        (36, 1, "e"),
        (36, 2, ""),  # not covered specifically => ''
        # cdi=9999 => not covered => returns ""
        (9999, 1, "")
    ]

    print("Running local unit tests for 'determine_type_of_com'...")
    all_passed = True
    for (cdi, cdn, expected) in test_cases:
        result = determine_type_of_com(cdi, cdn)
        if result != expected:
            print(f"  FAIL: determine_type_of_com({cdi}, {cdn}) = '{result}', expected '{expected}'")
            all_passed = False
        else:
            print(f"  PASS: ({cdi}, {cdn}) => '{result}'")

    if all_passed:
        print("All local unit tests PASSED.")
    else:
        print("Some local unit tests FAILED.")

#
# PART 2: Integration Tests (with real PM data)
#

def get_credentials():
    """
    Load login credentials from config.json or any place you prefer.
    """
    with open("config.json") as f:
        cfg = json.load(f)
    return cfg["USERNAME"], cfg["PASSWORD"]

def find_example_for_cdi_cdn(session, cdi, cdn, start_date=None, days_to_check=90):
    """
    Scans backward from `start_date` (default= today) up to `days_to_check` days,
    looking for actual data for (cdi, cdn). Returns the first date_str found, or None if not found.
    """
    if not start_date:
        start_date = datetime.date.today()

    # We go backwards day by day
    for offset in range(days_to_check):
        day = start_date - datetime.timedelta(days=offset)
        date_str = day.strftime("%m/%d/%Y")
        try:
            rows = fetch_activity_detail(session, date_str, cdi, cdn)
            if rows:  # if not empty, we found real data
                return date_str
        except Exception as exc:
            # e.g. some 404 or other error
            print(f"Skipping {date_str} due to error: {exc}")
    return None

def run_integration_tests():
    """
    For each (cdi, cdn), tries to find a real example day in the last N days
    and compares the parse result's 'method' to what 'determine_type_of_com' says.
    """
    username, password = get_credentials()
    session = login_to_practice_mojo(username, password)

    # Basic guess: method_name => 'Email' => we expect 'e', 'Text' => 't', 'Letter' => 'l'.
    # But your real data might name them differently. Adjust as needed:
    method_map = {
        "Email": "e",
        "Text Message": "t",
        "Postcard": "l"
    }

    # We'll test a selection of cdi values, and cdn in [1,2,3].
    # Add or remove as needed:
    cdi_list = [1, 8, 21, 22, 23, 30, 33, 35, 36, 130]
    cdn_list = [1, 2, 3]

    print("Running integration tests with real PracticeMojo data...\n")
    for cdi in cdi_list:
        for cdn in cdn_list:
            type_of_com_expected = determine_type_of_com(cdi, cdn)  # Our function's guess
            # find a real day that has data for (cdi, cdn)
            date_found = find_example_for_cdi_cdn(session, cdi, cdn, days_to_check=60)
            if not date_found:
                print(f"[CDI={cdi}, CDN={cdn}] No data found in last 60 days.")
                continue

            rows = fetch_activity_detail(session, date_found, cdi, cdn)
            # We'll just take the first row we see:
            row = rows[0]
            actual_method = row.get("method", "Unknown")
            
            # If actual_method is "Email"/"Text"/"Letter", we compare to our function's letter
            if actual_method in method_map:
                real_letter = method_map[actual_method]
                if real_letter == type_of_com_expected:
                    print(f"[CDI={cdi}, CDN={cdn}, Date={date_found}] PASS => PM says '{actual_method}', function says '{type_of_com_expected}'.")
                else:
                    print(f"[CDI={cdi}, CDN={cdn}, Date={date_found}] FAIL => PM says '{actual_method}' => '{real_letter}', function says '{type_of_com_expected}'.")
            else:
                print(f"[CDI={cdi}, CDN={cdn}, Date={date_found}] Unknown method='{actual_method}'. (Adjust 'method_map' if needed.)")

    print("\nIntegration tests complete.")

if __name__ == "__main__":
    run_unit_tests()
    print()
    run_integration_tests()
