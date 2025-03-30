import requests
from bs4 import BeautifulSoup
import json
import sys

def login_to_practice_mojo(login_id, password):
    """
    Logs into PracticeMojo and returns a requests.Session() that is authenticated.
    """
    session = requests.Session()
    
    # Optional: GET the login page first for cookies.
    login_page_url = "https://app.practicemojo.com/Pages/login"
    session.get(login_page_url)

    # POST to the login endpoint
    login_url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/userLogin"
    form_data = {
        "loginId": login_id,
        "password": password,
        "slug": "login"
    }
    headers = {
        "Referer": login_page_url,
        "Content-Type": "application/x-www-form-urlencoded"
    }
    response = session.post(login_url, data=form_data, headers=headers)
    
    if response.status_code != 200:
        raise Exception(f"Login failed with status code {response.status_code}.")
    
    # Check if the response has "Invalid login" or similar
    if "Invalid login" in response.text or "Incorrect password" in response.text:
        raise Exception("Login failed: invalid credentials or site error.")
    
    return session

def fetch_activity_detail(session, date_str="03/29/2025", cdi=130, cdn=2):
    """
    GET the detail page for a given date/cdi/cdn, return a BeautifulSoup object.
    """
    detail_url = (
        "https://app.practicemojo.com/"
        f"cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?"
        f"td={date_str}&cdi={cdi}&cdn={cdn}"
    )
    response = session.get(detail_url)
    
    if response.status_code != 200:
        raise Exception(f"Failed to fetch detail page: status={response.status_code}")
    
    # Parse HTML
    soup = BeautifulSoup(response.text, "html.parser")
    return soup

def parse_detail_page(soup):
    """
    Extract rows from the 'Activity Detail' table.
    Returns a list of dicts with:
      campaign, method, patient_name, appointment, confirmations, status
    """
    results = []
    
    # 1) Locate the main table inside .activity_detail_table
    table = soup.select_one("div.activity_detail_table table")
    if not table:
        print("[parse_detail_page] No table found in .activity_detail_table!", file=sys.stderr)
        return results
    
    # 2) Because the top-level <tr> are inside <tbody> or otherwise, 
    #    sometimes we do or do not need `recursive=False`. We'll just grab all <tr>.
    rows = table.find_all("tr")
    
    i = 0
    while i < len(rows):
        tds = rows[i].find_all("td")
        
        # Check if this row starts a new "campaign" block:
        # The snippet typically has tds[0] and tds[1] with rowspans, plus an extra "colspan=4".
        if len(tds) >= 3 and tds[0].has_attr("rowspan") and tds[1].has_attr("rowspan"):
            campaign_name = tds[0].get_text(strip=True)
            method_name = tds[1].get_text(strip=True)
            
            sub_count = int(tds[0]["rowspan"])
            
            # Advance to the next row, which has actual patient data
            i += 1
            block_end = i + (sub_count - 1)
            
            while i < block_end and i < len(rows):
                data_tds = rows[i].find_all("td")
                if len(data_tds) == 4:
                    patient_name = data_tds[0].get_text(strip=True)
                    appointment  = data_tds[1].get_text(strip=True)
                    confirmations = data_tds[2].get_text(strip=True)
                    status       = data_tds[3].get_text(strip=True)
                    
                    # Remove the arrow if present
                    # or just do a replace on arrow + any whitespace
                    patient_name = patient_name.replace("↳", "").strip()
                    
                    # Optionally skip "Family" lines if you only want the actual individuals:
                    if "Family" in patient_name:
                        # skip
                        pass
                    else:
                        row_dict = {
                            "campaign": campaign_name,
                            "method": method_name,
                            "patient_name": patient_name,
                            "appointment": appointment,
                            "confirmations": confirmations,
                            "status": status
                        }
                        results.append(row_dict)
                i += 1
        else:
            i += 1
    
    return results

def main():
    # Load credentials
    with open('config.json') as f:
        config = json.load(f)

    login_id = config.get('USERNAME')
    password = config.get('PASSWORD')
    
    # 1) Log in
    session = login_to_practice_mojo(login_id, password)
    print("Logged in successfully!")
    
    # 2) Fetch detail page
    date_str = "03/29/2025"
    cdi = 130
    cdn = 2
    soup = fetch_activity_detail(session, date_str, cdi, cdn)
    
    # 3) Parse
    data_rows = parse_detail_page(soup)
    
    # 4) Print results
    print("Extracted data rows:")
    for row in data_rows:
        print(row)

if __name__ == "__main__":
    main()
