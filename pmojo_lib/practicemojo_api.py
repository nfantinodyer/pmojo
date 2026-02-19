"""PracticeMojo API: login and data fetching via HTTP requests."""
import requests
from bs4 import BeautifulSoup


class PracticeMojoAPI:
    """Handles PracticeMojo login and fetching campaign activity data."""

    BASE_URL = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/"

    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.session = None

    def login(self):
        """Logs in and sets self.session to an authenticated requests.Session."""
        s = requests.Session()
        login_page_url = "https://app.practicemojo.com/Pages/login"
        s.get(login_page_url)

        login_url = self.BASE_URL + "userLogin"
        form_data = {
            "loginId": self.username,
            "password": self.password,
            "slug": "login"
        }
        headers = {
            "Referer": login_page_url,
            "Content-Type": "application/x-www-form-urlencoded"
        }
        resp = s.post(login_url, data=form_data, headers=headers)
        if resp.status_code != 200:
            raise Exception(f"Login failed with status code {resp.status_code}")
        if "Invalid login" in resp.text or "Incorrect password" in resp.text:
            raise Exception("Login failed: invalid credentials or site error.")

        self.session = s

    def fetch_activity_detail(self, date_str: str, cdi: int, cdn: int):
        """Fetch the detail page for (date_str, cdi, cdn), parse and return row dicts."""
        if not self.session:
            raise Exception("Not logged in to PracticeMojo.")
        detail_url = f"{self.BASE_URL}gotoActivityDetail?td={date_str}&cdi={cdi}&cdn={cdn}"
        resp = self.session.get(detail_url)
        if resp.status_code != 200:
            raise Exception(f"Failed to fetch detail page (cdi={cdi}, cdn={cdn}): status={resp.status_code}")
        soup = BeautifulSoup(resp.text, "lxml")
        return self.parse_activity_detail(soup)

    @staticmethod
    def parse_activity_detail(soup: BeautifulSoup):
        """
        Parse the HTML soup for the campaign detail table and return a list of dicts.

        Supports both the 2026+ structure (div.table-responsive) and the
        legacy structure (div.activity_detail_table with rowspan).
        """
        results = []

        # Try new structure first, fall back to legacy
        table = soup.select_one("div.table-responsive table")
        if table:
            return PracticeMojoAPI._parse_new_structure(table)

        # Legacy fallback (pre-2026 UI)
        table = soup.select_one("div.activity_detail_table table")
        if table:
            return PracticeMojoAPI._parse_legacy_structure(table)

        return results

    @staticmethod
    def _parse_new_structure(table):
        """Parse the 2026+ PracticeMojo HTML table structure."""
        results = []
        tbody = table.find("tbody")
        if not tbody:
            return results

        rows = tbody.find_all("tr")
        campaign_name = ""
        method_name = ""

        for row in rows:
            cls = row.get("class", [])

            # Group header row: has campaign name and method
            if "table-secondary" in cls:
                camp_div = row.find("div", class_="fw-bold text-primary")
                if camp_div:
                    raw_campaign = camp_div.get_text(strip=True)
                    # Strip the count suffix e.g. "Recare: Due - (1)" -> "Recare: Due"
                    idx = raw_campaign.rfind(" - (")
                    campaign_name = raw_campaign[:idx] if idx != -1 else raw_campaign

                # Method is the fw-bold div that does NOT have text-primary
                bold_divs = row.find_all("div", class_="fw-bold")
                for div in bold_divs:
                    if "text-primary" not in (div.get("class") or []):
                        method_name = div.get_text(strip=True)
                        break
                continue

            # Family row: skip
            if "table-light" in cls:
                continue

            # Patient row: has a link to gotoPatientDetail
            link = row.find("a", href=lambda h: h and "gotoPatientDetail" in h)
            if not link:
                continue

            patient_name = link.get_text(strip=True)
            if "Family" in patient_name:
                continue

            # Extract data from the 6 TDs (many may be empty)
            tds = row.find_all("td")
            appointment = tds[3].get_text(strip=True) if len(tds) > 3 else ""
            confirmations = tds[4].get_text(strip=True) if len(tds) > 4 else ""
            row_status = tds[5].get_text(strip=True) if len(tds) > 5 else ""

            results.append({
                "campaign": campaign_name,
                "method": method_name,
                "patient_name": patient_name,
                "appointment": appointment,
                "confirmations": confirmations,
                "status": row_status
            })

        return results

    @staticmethod
    def _parse_legacy_structure(table):
        """Parse the pre-2026 PracticeMojo HTML table structure (rowspan-based)."""
        results = []
        rows = table.find_all("tr", recursive=False) or table.find_all("tr")
        i = 0
        while i < len(rows):
            tds = rows[i].find_all("td", recursive=False)
            if len(tds) >= 3 and tds[0].has_attr("rowspan") and tds[1].has_attr("rowspan"):
                campaign_name = tds[0].get_text(strip=True)
                method_name = tds[1].get_text(strip=True)
                sub_count = int(tds[0]["rowspan"])
                i += 1
                block_end = i + (sub_count - 1)

                while i < block_end and i < len(rows):
                    data_row = rows[i]
                    data_tds = data_row.find_all("td", recursive=False)
                    if len(data_tds) == 4:
                        patient_name = data_tds[0].get_text().replace("\u21b3", "").rstrip()
                        if "Family" not in patient_name:
                            appointment = data_tds[1].get_text(strip=True)
                            confirmations = data_tds[2].get_text(strip=True)
                            row_status = data_tds[3].get_text(strip=True)
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
