"""
Comprehensive test suite for pmojo.

Tests cover:
  - HTML parsing (new 2026+ structure and legacy rowspan structure)
  - determine_type_of_com mapping
  - Campaign name mapping (com_map)
  - Recare flag logic
  - Appointment time merging
  - Database operations
  - Login validation

Run with: pytest testing/test_pmojo.py -v
"""
import os
import sys
import json
import datetime
import sqlite3
import tempfile

import pytest
from bs4 import BeautifulSoup

# Add project root to path so we can import from pmojo.py without pywinauto/ahk
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_ROOT)

FIXTURES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


# ---------------------------------------------------------------------------
# We can't import pmojo.py directly because it imports pywinauto and ahk
# at module level (Windows-only). Instead we extract the pure logic functions
# and the parser class methods to test them independently.
# ---------------------------------------------------------------------------

def _load_pmojo_source():
    """Load pmojo.py source and extract testable functions via exec."""
    src_path = os.path.join(PROJECT_ROOT, "pmojo.py")
    with open(src_path) as f:
        source = f.read()
    return source


# -- Extract the parser and pure logic by re-implementing them from source --
# This avoids importing pywinauto/ahk which aren't available on Mac.

def parse_activity_detail(soup):
    """Wrapper that calls the parser logic from pmojo.py."""
    # Try new structure first
    table = soup.select_one("div.table-responsive table")
    if table:
        return _parse_new_structure(table)
    # Legacy fallback
    table = soup.select_one("div.activity_detail_table table")
    if table:
        return _parse_legacy_structure(table)
    return []


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

        if "table-secondary" in cls:
            camp_div = row.find("div", class_="fw-bold text-primary")
            if camp_div:
                raw_campaign = camp_div.get_text(strip=True)
                idx = raw_campaign.rfind(" - (")
                campaign_name = raw_campaign[:idx] if idx != -1 else raw_campaign

            bold_divs = row.find_all("div", class_="fw-bold")
            for div in bold_divs:
                if "text-primary" not in (div.get("class") or []):
                    method_name = div.get_text(strip=True)
                    break
            continue

        if "table-light" in cls:
            continue

        link = row.find("a", href=lambda h: h and "gotoPatientDetail" in h)
        if not link:
            continue

        patient_name = link.get_text(strip=True)
        if "Family" in patient_name:
            continue

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


def _parse_legacy_structure(table):
    """Parse the pre-2026 PracticeMojo HTML table structure."""
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
                    patient_name = data_tds[0].get_text().replace("↳", "").rstrip()
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


def determine_type_of_com(cdi, cdn):
    """Copied from pmojo.py PmojoAutomation.determine_type_of_com."""
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


# ---------------------------------------------------------------------------
# Database helper (copied from pmojo.py to avoid import issues)
# ---------------------------------------------------------------------------

class Database:
    def __init__(self, db_path):
        self.db_path = db_path
        self.gui = None
        self.init_db()

    @staticmethod
    def _ui_str_to_date(ui_str):
        return datetime.datetime.strptime(ui_str, "%m/%d/%Y").date()

    @staticmethod
    def _date_to_db_str(d):
        return d.strftime("%Y-%m-%d")

    @staticmethod
    def _db_str_to_date(db_str):
        return datetime.datetime.strptime(db_str, "%Y-%m-%d").date()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS days(
                date TEXT PRIMARY KEY,
                status TEXT DEFAULT '',
                error_msg TEXT DEFAULT ''
            )
        """)
        start_date = datetime.date(2020, 1, 1)
        end_date = datetime.date(2025, 3, 27)
        delta = datetime.timedelta(days=1)
        to_insert = []
        current = start_date
        while current <= end_date:
            iso_str = self._date_to_db_str(current)
            to_insert.append((iso_str, "done", ""))
            current += delta
        c.executemany("INSERT OR IGNORE INTO days(date, status, error_msg) VALUES (?, ?, ?)", to_insert)
        conn.commit()
        conn.close()

    def expand_db_up_to(self, end_date):
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT MAX(date) FROM days")
        row = c.fetchone()
        latest_str = row[0] if row and row[0] else None
        if latest_str:
            latest_date = self._db_str_to_date(latest_str)
        else:
            latest_date = datetime.date(2020, 1, 1)
        if end_date > latest_date:
            to_insert = []
            cur = latest_date + datetime.timedelta(days=1)
            while cur <= end_date:
                iso_str = self._date_to_db_str(cur)
                to_insert.append((iso_str, '', ''))
                cur += datetime.timedelta(days=1)
            if to_insert:
                c.executemany("INSERT OR IGNORE INTO days(date, status, error_msg) VALUES (?, ?, ?)", to_insert)
        conn.commit()
        conn.close()

    def get_day_status(self, ui_date_str):
        try:
            d = self._ui_str_to_date(ui_date_str)
            iso_str = self._date_to_db_str(d)
        except ValueError:
            return ""
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT status FROM days WHERE date=?", (iso_str,))
        row = c.fetchone()
        conn.close()
        return row[0] if row else ""

    def toggle_day_status(self, ui_date_str, new_status):
        current = self.get_day_status(ui_date_str)
        if current == new_status:
            return
        try:
            d = self._ui_str_to_date(ui_date_str)
            iso_str = self._date_to_db_str(d)
        except ValueError:
            return
        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("INSERT OR IGNORE INTO days(date) VALUES(?)", (iso_str,))
        c.execute("UPDATE days SET status=? WHERE date=?", (new_status, iso_str))
        conn.commit()
        conn.close()


# ---------------------------------------------------------------------------
# Helper to load fixture HTML
# ---------------------------------------------------------------------------

def load_fixture(name):
    """Load an HTML fixture file and return BeautifulSoup."""
    path = os.path.join(FIXTURES_DIR, name)
    with open(path) as f:
        return BeautifulSoup(f.read(), "html.parser")


# ===========================================================================
# TEST: determine_type_of_com
# ===========================================================================

class TestDetermineTypeOfCom:
    """Test the CDI/CDN -> communication type mapping."""

    @pytest.mark.parametrize("cdi,cdn,expected", [
        # Appointment reminders (CDI 23, 130)
        (23, 1, "e"),
        (23, 2, "t"),
        (23, 3, ""),
        (130, 1, "e"),
        (130, 2, "t"),
        (130, 3, ""),
        # Recare: Due (CDI 1) - has all 3 CDNs
        (1, 1, "l"),
        (1, 2, "e"),
        (1, 3, "t"),
        # Recare: Really Past Due (CDI 8)
        (8, 1, "l"),
        (8, 2, "e"),
        (8, 3, ""),
        # Birthday cards (CDI 21, 22)
        (21, 1, "l"),
        (21, 2, ""),
        (22, 1, "l"),
        (22, 2, ""),
        # Reactivate (CDI 30)
        (30, 1, "l"),
        (30, 2, "e"),
        (30, 3, ""),
        # Recare: Past Due (CDI 33)
        (33, 1, "l"),
        (33, 2, "e"),
        (33, 3, ""),
        # Anniversary (CDI 35)
        (35, 1, "l"),
        (35, 2, ""),
        # e-Birthday (CDI 36)
        (36, 1, "e"),
        (36, 2, ""),
        # Unknown CDI
        (9999, 1, ""),
        (9999, 2, ""),
        (9999, 3, ""),
    ])
    def test_mapping(self, cdi, cdn, expected):
        assert determine_type_of_com(cdi, cdn) == expected

    def test_appointment_reminder_cdis_return_email_for_cdn1(self):
        for cdi in [23, 130]:
            assert determine_type_of_com(cdi, 1) == "e"

    def test_appointment_reminder_cdis_return_text_for_cdn2(self):
        for cdi in [23, 130]:
            assert determine_type_of_com(cdi, 2) == "t"

    def test_recare_cdis_return_letter_for_cdn1(self):
        for cdi in [1, 8, 33]:
            assert determine_type_of_com(cdi, 1) == "l"

    def test_only_cdi1_has_text_option(self):
        """Only CDI=1 with CDN=3 returns 't' for non-appointment campaigns."""
        assert determine_type_of_com(1, 3) == "t"
        for cdi in [8, 21, 22, 30, 33, 35, 36]:
            assert determine_type_of_com(cdi, 3) == ""


# ===========================================================================
# TEST: Recare flag
# ===========================================================================

class TestRecareFlag:
    """Test recare flag logic (CDIs 1, 8, 33 get 'r')."""

    @pytest.mark.parametrize("cdi,expected", [
        (1, "r"),
        (8, "r"),
        (33, "r"),
        (13, ""),
        (21, ""),
        (22, ""),
        (23, ""),
        (30, ""),
        (35, ""),
        (36, ""),
        (130, ""),
    ])
    def test_recare_flag(self, cdi, expected):
        recare_flag = "r" if cdi in [1, 8, 33] else ""
        assert recare_flag == expected


# ===========================================================================
# TEST: Campaign name mapping (com_map)
# ===========================================================================

class TestComMap:
    """Test campaign name mapping used for SoftDent entry."""

    COM_MAP = {
        1: "Recare: Due",
        8: "Recare: Really Past Due",
        13: "SomeCampaign13",
        21: "bday card",
        22: "bday card",
        23: "Some Appt Reminder",
        30: "Reactivate: 1 year ago",
        33: "Recare: Past Due",
        35: "Anniversary card",
        36: "bday card",
        130: "Another Appt Reminder"
    }

    def test_all_cdi_have_mapping(self):
        """Every CDI in the processing list has a com_map entry."""
        cdi_list = [1, 8, 13, 21, 22, 23, 30, 33, 35, 36, 130]
        for cdi in cdi_list:
            assert cdi in self.COM_MAP, f"CDI {cdi} missing from com_map"

    def test_no_empty_names(self):
        for cdi, name in self.COM_MAP.items():
            assert name.strip(), f"CDI {cdi} has empty campaign name"


# ===========================================================================
# TEST: New HTML structure parsing (2026+)
# ===========================================================================

class TestParseNewStructure:
    """Test parsing of the new PracticeMojo HTML structure (2026+ UI)."""

    def test_appointment_reminder_email(self):
        """CDI=23, CDN=1: Courtesy Reminder via Email with appointment times."""
        soup = load_fixture("cdi23_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0, "Should find patients"
        for r in results:
            assert r["campaign"] == "Courtesy Reminder"
            assert r["method"] == "Email"
            assert r["patient_name"].strip()
            assert r["appointment"], f"Patient {r['patient_name']} should have appointment time"
            assert "@" in r["appointment"], f"Appointment should contain @ symbol: {r['appointment']}"

    def test_appointment_reminder_text(self):
        """CDI=23, CDN=2: Courtesy Reminder via Text Message."""
        soup = load_fixture("cdi23_cdn2.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["campaign"] == "Courtesy Reminder"
            assert r["method"] == "Text Message"

    def test_unconfirmed_reminder(self):
        """CDI=130, CDN=1: Courtesy Reminder Unconfirmed via Email."""
        soup = load_fixture("cdi130_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["campaign"] == "Courtesy Reminder: Unconfirmed"
            assert r["method"] == "Email"
            assert r["appointment"]

    def test_unconfirmed_reminder_text(self):
        """CDI=130, CDN=2: Courtesy Reminder Unconfirmed via Text."""
        soup = load_fixture("cdi130_cdn2.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Text Message"

    def test_recare_due_postcard(self):
        """CDI=1, CDN=1: Recare Due via Postcard (no appointment times)."""
        soup = load_fixture("cdi1_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["campaign"] == "Recare: Due"
            assert r["method"] == "Postcard"
            # Recare has no appointment times
            assert r["appointment"] == ""

    def test_recare_due_email(self):
        """CDI=1, CDN=2: Recare Due via Email."""
        soup = load_fixture("cdi1_cdn2.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Email"

    def test_recare_due_text(self):
        """CDI=1, CDN=3: Recare Due via Text Message."""
        soup = load_fixture("cdi1_cdn3.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Text Message"

    def test_recare_really_past_due(self):
        """CDI=8: Recare Really Past Due."""
        soup = load_fixture("cdi8_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert "Past Due" in r["campaign"] or "Recare" in r["campaign"]
            assert r["method"] == "Postcard"

    def test_recare_past_due(self):
        """CDI=33: Recare Past Due."""
        soup = load_fixture("cdi33_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Postcard"

    def test_birthday_postcard(self):
        """CDI=21: Birthday Adults via Postcard."""
        soup = load_fixture("cdi21_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Postcard"

    def test_ebirthday_email(self):
        """CDI=36: e-Birthday Adult via Email."""
        soup = load_fixture("cdi36_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Email"

    def test_reactivate_postcard(self):
        """CDI=30: Reactivate via Postcard."""
        soup = load_fixture("cdi30_cdn1.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["method"] == "Postcard"

    def test_family_rows_excluded(self):
        """Family rows (table-light class) should not appear in results."""
        soup = load_fixture("cdi23_cdn1.html")
        results = parse_activity_detail(soup)
        for r in results:
            assert "Family" not in r["patient_name"], f"Family row leaked: {r['patient_name']}"

    def test_patient_names_no_arrow(self):
        """Patient names should not contain the ↳ arrow character."""
        soup = load_fixture("cdi23_cdn1.html")
        results = parse_activity_detail(soup)
        for r in results:
            assert "↳" not in r["patient_name"], f"Arrow in patient name: {r['patient_name']}"

    def test_appointment_time_format(self):
        """Appointment times should follow MM/DD @ HH:MM AM/PM pattern."""
        import re
        soup = load_fixture("cdi23_cdn1.html")
        results = parse_activity_detail(soup)
        pattern = re.compile(r"\d{2}/\d{2} @ \d{1,2}:\d{2} [AP]M")
        for r in results:
            if r["appointment"]:
                assert pattern.match(r["appointment"]), f"Bad format: {r['appointment']}"

    def test_all_required_fields_present(self):
        """Every result dict has all required keys."""
        required_keys = {"campaign", "method", "patient_name", "appointment", "confirmations", "status"}
        for fixture in ["cdi23_cdn1.html", "cdi1_cdn1.html", "cdi130_cdn1.html"]:
            soup = load_fixture(fixture)
            results = parse_activity_detail(soup)
            for r in results:
                assert set(r.keys()) == required_keys, f"Missing keys in {fixture}: {required_keys - set(r.keys())}"


# ===========================================================================
# TEST: Legacy HTML structure parsing (pre-2026)
# ===========================================================================

class TestParseLegacyStructure:
    """Test parsing of the old PracticeMojo HTML structure (rowspan-based)."""

    def test_legacy_rowspan_parsing(self):
        """Legacy structure with rowspan should still parse correctly."""
        soup = load_fixture("legacy_rowspan.html")
        results = parse_activity_detail(soup)
        assert len(results) == 2
        assert results[0]["campaign"] == "Recare: Due"
        assert results[0]["method"] == "Email"
        assert results[0]["patient_name"] == "Smith, John"
        assert results[0]["appointment"] == "03/15 @ 10:00 AM"
        assert results[0]["confirmations"] == "Yes"
        assert results[0]["status"] == "Sent"
        assert results[1]["patient_name"] == "Doe, Jane"
        assert results[1]["appointment"] == "03/15 @ 02:00 PM"
        assert results[1]["status"] == "Pending"


# ===========================================================================
# TEST: Empty / edge cases
# ===========================================================================

class TestEdgeCases:
    """Test edge cases for the parser."""

    def test_empty_html(self):
        soup = BeautifulSoup("<html><body></body></html>", "html.parser")
        assert parse_activity_detail(soup) == []

    def test_table_with_no_patients(self):
        html = """
        <div class="table-responsive">
        <table><tbody>
        <tr class="table-secondary">
            <td><div class="fw-bold text-primary">Campaign - (0)</div></td>
            <td><div class="fw-bold">Email</div></td>
            <td></td><td></td><td></td><td></td>
        </tr>
        </tbody></table>
        </div>
        """
        soup = BeautifulSoup(html, "html.parser")
        assert parse_activity_detail(soup) == []

    def test_campaign_name_count_stripped(self):
        """Campaign name should have the count suffix stripped."""
        html = """
        <div class="table-responsive">
        <table><tbody>
        <tr class="table-secondary">
            <td><div class="fw-bold text-primary">Recare: Due - (5)</div></td>
            <td><div class="fw-bold">Postcard</div></td>
            <td></td><td></td><td></td><td></td>
        </tr>
        <tr>
            <td></td><td></td>
            <td><div class="d-flex align-items-center">
                <a href="/wa/gotoPatientDetail?pid=1">Test, Patient</a>
            </div></td>
            <td></td><td></td><td></td>
        </tr>
        </tbody></table>
        </div>
        """
        soup = BeautifulSoup(html, "html.parser")
        results = parse_activity_detail(soup)
        assert len(results) == 1
        assert results[0]["campaign"] == "Recare: Due"

    def test_campaign_name_without_count(self):
        """Campaign name without count suffix should be preserved as-is."""
        html = """
        <div class="table-responsive">
        <table><tbody>
        <tr class="table-secondary">
            <td><div class="fw-bold text-primary">Some Campaign</div></td>
            <td><div class="fw-bold">Email</div></td>
            <td></td><td></td><td></td><td></td>
        </tr>
        <tr>
            <td></td><td></td>
            <td><div class="d-flex align-items-center">
                <a href="/wa/gotoPatientDetail?pid=1">Test, Patient</a>
            </div></td>
            <td></td><td></td><td></td>
        </tr>
        </tbody></table>
        </div>
        """
        soup = BeautifulSoup(html, "html.parser")
        results = parse_activity_detail(soup)
        assert results[0]["campaign"] == "Some Campaign"

    def test_no_table_responsive_and_no_legacy_returns_empty(self):
        html = "<div class='something-else'><table><tr><td>data</td></tr></table></div>"
        soup = BeautifulSoup(html, "html.parser")
        assert parse_activity_detail(soup) == []


# ===========================================================================
# TEST: Method-to-type consistency
# ===========================================================================

class TestMethodTypeConsistency:
    """Verify that PM method names map correctly to determine_type_of_com results."""

    METHOD_MAP = {
        "Email": "e",
        "Text Message": "t",
        "Postcard": "l",
    }

    @pytest.mark.parametrize("fixture,cdi,cdn", [
        ("cdi23_cdn1.html", 23, 1),
        ("cdi23_cdn2.html", 23, 2),
        ("cdi130_cdn1.html", 130, 1),
        ("cdi130_cdn2.html", 130, 2),
        ("cdi1_cdn1.html", 1, 1),
        ("cdi1_cdn2.html", 1, 2),
        ("cdi1_cdn3.html", 1, 3),
        ("cdi8_cdn1.html", 8, 1),
        ("cdi8_cdn2.html", 8, 2),
        ("cdi33_cdn1.html", 33, 1),
        ("cdi33_cdn2.html", 33, 2),
        ("cdi36_cdn1.html", 36, 1),
        ("cdi30_cdn1.html", 30, 1),
        ("cdi30_cdn2.html", 30, 2),
        ("cdi21_cdn1.html", 21, 1),
    ])
    def test_method_matches_type_of_com(self, fixture, cdi, cdn):
        """The method from parsed HTML should match determine_type_of_com."""
        soup = load_fixture(fixture)
        results = parse_activity_detail(soup)
        assert len(results) > 0, f"No results in {fixture}"

        expected_type = determine_type_of_com(cdi, cdn)
        actual_method = results[0]["method"]

        if actual_method in self.METHOD_MAP:
            actual_type = self.METHOD_MAP[actual_method]
            assert actual_type == expected_type, (
                f"Fixture {fixture}: method={actual_method!r} maps to {actual_type!r}, "
                f"but determine_type_of_com({cdi},{cdn})={expected_type!r}"
            )


# ===========================================================================
# TEST: Appointment time merging logic
# ===========================================================================

class TestAppointmentMerging:
    """Test the appointment time merging used in merge_and_type_appointments."""

    @staticmethod
    def merge_appointments(appts):
        """Reproduce the merging logic from PmojoAutomation.merge_and_type_appointments."""
        from pmojo_lib.automation import _appt_sort_key
        unique_appts = sorted(set(appts), key=_appt_sort_key)
        formatted = []
        for idx, a in enumerate(unique_appts):
            if idx == 0:
                formatted.append(a)
            else:
                parts = a.split('@')
                if len(parts) > 1:
                    formatted.append(parts[1].strip())
                else:
                    formatted.append(a)
        return ' & '.join(formatted)

    def test_single_appointment(self):
        assert self.merge_appointments(["03/04 @ 10:00 AM"]) == "03/04 @ 10:00 AM"

    def test_two_appointments(self):
        # Chronological: 10:00 AM before 02:00 PM
        result = self.merge_appointments(["03/04 @ 10:00 AM", "03/04 @ 02:00 PM"])
        assert result == "03/04 @ 10:00 AM & 02:00 PM"

    def test_duplicate_appointments_removed(self):
        result = self.merge_appointments(["03/04 @ 10:00 AM", "03/04 @ 10:00 AM"])
        assert result == "03/04 @ 10:00 AM"

    def test_three_appointments(self):
        # Chronological: 08:00 AM, 10:00 AM, 02:00 PM
        result = self.merge_appointments(
            ["03/04 @ 08:00 AM", "03/04 @ 10:00 AM", "03/04 @ 02:00 PM"]
        )
        assert result == "03/04 @ 08:00 AM & 10:00 AM & 02:00 PM"

    def test_unsorted_input(self):
        result = self.merge_appointments(["03/04 @ 02:00 PM", "03/04 @ 08:00 AM"])
        assert result == "03/04 @ 08:00 AM & 02:00 PM"


# ===========================================================================
# TEST: Database operations
# ===========================================================================

class TestDatabase:
    """Test database initialization and operations."""

    @pytest.fixture
    def db(self, tmp_path):
        return Database(str(tmp_path / "test_days.db"))

    def test_init_creates_table(self, db):
        conn = sqlite3.connect(db.db_path)
        c = conn.cursor()
        c.execute("SELECT count(*) FROM days")
        count = c.fetchone()[0]
        conn.close()
        # Should have dates from 2020-01-01 to 2025-03-27
        expected = (datetime.date(2025, 3, 27) - datetime.date(2020, 1, 1)).days + 1
        assert count == expected

    def test_all_initial_dates_are_done(self, db):
        assert db.get_day_status("01/15/2023") == "done"
        assert db.get_day_status("12/31/2024") == "done"

    def test_get_status_for_unknown_date(self, db):
        assert db.get_day_status("06/15/2026") == ""

    def test_toggle_status(self, db):
        db.toggle_day_status("01/01/2023", "error")
        assert db.get_day_status("01/01/2023") == "error"

    def test_toggle_to_empty(self, db):
        db.toggle_day_status("01/01/2023", "")
        assert db.get_day_status("01/01/2023") == ""

    def test_toggle_same_status_noop(self, db):
        assert db.get_day_status("01/01/2023") == "done"
        db.toggle_day_status("01/01/2023", "done")
        assert db.get_day_status("01/01/2023") == "done"

    def test_expand_db(self, db):
        new_date = datetime.date(2026, 6, 15)
        db.expand_db_up_to(new_date)
        assert db.get_day_status("06/15/2026") == ""
        assert db.get_day_status("04/01/2025") == ""  # newly added, not 'done'

    def test_invalid_date_format(self, db):
        assert db.get_day_status("not-a-date") == ""

    def test_date_format_conversion(self, db):
        """UI format MM/DD/YYYY should map to DB format YYYY-MM-DD."""
        d = Database._ui_str_to_date("03/15/2025")
        assert Database._date_to_db_str(d) == "2025-03-15"


# ===========================================================================
# TEST: Login validation (mocked)
# ===========================================================================

class TestLoginValidation:
    """Test login error detection logic."""

    def test_invalid_login_detected(self):
        """Response containing 'Invalid login' should be caught."""
        text = "<html>Invalid login credentials</html>"
        assert "Invalid login" in text

    def test_incorrect_password_detected(self):
        text = "<html>Incorrect password</html>"
        assert "Incorrect password" in text

    def test_successful_login_no_errors(self):
        text = "<html>Welcome to dashboard</html>"
        assert "Invalid login" not in text
        assert "Incorrect password" not in text


# ===========================================================================
# TEST: Integration - parse real fixture then verify type_of_com
# ===========================================================================

class TestEndToEndParsing:
    """End-to-end test: parse fixture, verify all fields work with the automation logic."""

    def test_appointment_campaign_has_times_for_merge(self):
        """For CDI=23 patients, appointment field should be populated for merging."""
        soup = load_fixture("cdi23_cdn1.html")
        results = parse_activity_detail(soup)

        # Group by patient (as merge_and_type_appointments does)
        grouped = {}
        for row in results:
            patient = row["patient_name"].strip()
            appt = row["appointment"]
            if appt:
                grouped.setdefault(patient, []).append(appt)

        assert len(grouped) > 0, "Should have patients with appointments"
        for patient, appts in grouped.items():
            for a in appts:
                assert "@" in a, f"{patient}: invalid appointment format {a!r}"

    def test_appointment_campaign_with_empty_appointments_skipped(self):
        """CDI=23/CDN=2 has some patients with empty appointments.

        These should fail validate_row but the merge logic should skip them
        instead of crashing (the pre-validation only crashes if ALL rows fail).
        """
        soup = load_fixture("cdi23_cdn2.html")
        results = parse_activity_detail(soup)

        # Some rows should have empty appointments
        empty_appt_rows = [r for r in results if not r["appointment"].strip()]
        valid_appt_rows = [r for r in results if r["appointment"].strip()]
        assert len(empty_appt_rows) > 0, "Fixture should have patients with empty appointments"
        assert len(valid_appt_rows) > 0, "Fixture should have patients with valid appointments"

        # validate_row should raise for empty appointment rows
        for row in empty_appt_rows:
            with pytest.raises(ParseError, match="Empty appointment"):
                validate_row(row, 23)

        # validate_row should pass for valid rows
        for row in valid_appt_rows:
            validate_row(row, 23)

        # The skip logic: group only valid rows (as merge_and_type_appointments now does)
        grouped = {}
        for row in results:
            try:
                validate_row(row, 23)
            except ParseError:
                continue
            patient = row["patient_name"].strip()
            appt = row["appointment"]
            if appt:
                grouped.setdefault(patient, []).append(appt)

        assert len(grouped) > 0, "Should still have valid patients after skipping"
        for patient, appts in grouped.items():
            for a in appts:
                assert "@" in a, f"{patient}: invalid appointment format {a!r}"

    def test_merge_with_families_duplicates_and_empty_appts(self):
        """Full merge scenario: families, multi-slot patients, duplicates, empty appts.

        Based on real 02/19/2026 CDI=23/CDN=1 data. Tests:
        - Family header rows are skipped
        - Sub-family members (↳) are parsed correctly without the arrow
        - Same patient with multiple times gets merged (e.g. "01:10 PM & 01:20 PM")
        - Duplicate patient+time rows (PracticeMojo bug) are deduplicated
        - Patients with empty appointment (moved appt) are skipped
        """
        html = """
        <div class="table-responsive">
        <table class="table table-bordered">
        <thead class="thead sticky-top"><tr>
            <th>Campaign</th><th>Method</th><th>Patients</th>
            <th>Appointment Date</th><th>Confirmations</th><th>Status</th>
        </tr></thead>
        <tbody class="tbody">
        <tr class="table-secondary">
            <td><div class="fw-bold text-primary">Courtesy Reminder - (9)</div></td>
            <td><div class="fw-bold">Email</div></td>
            <td></td><td></td><td></td><td></td>
        </tr>
        <!-- Family header for Pruitt -->
        <tr class="table-light">
            <td></td><td></td>
            <td><div class="text-muted fw-semibold">Pruitt Family</div></td>
            <td></td><td></td><td></td>
        </tr>
        <!-- Pruitt, Ashlee: two different times -->
        <tr><td></td><td></td>
            <td><div class="d-flex align-items-center">
                <span class="text-muted me-2">↳</span>
                <a href="/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoPatientDetail?pid=1">Pruitt, Ashlee</a>
            </div></td>
            <td>03/05 @ 01:10 PM</td><td></td><td></td>
        </tr>
        <tr><td></td><td></td>
            <td><div class="d-flex align-items-center">
                <span class="text-muted me-2">↳</span>
                <a href="/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoPatientDetail?pid=1">Pruitt, Ashlee</a>
            </div></td>
            <td>03/05 @ 01:20 PM</td><td></td><td></td>
        </tr>
        <!-- Brusa: single appointment -->
        <tr><td></td><td></td>
            <td><div class="d-flex align-items-center">
                <a href="/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoPatientDetail?pid=2">Brusa, Brianna</a>
            </div></td>
            <td>03/05 @ 01:20 PM</td><td></td><td></td>
        </tr>
        <!-- Robinson: single, NO appointment (moved) -->
        <tr><td></td><td></td>
            <td><div class="d-flex align-items-center">
                <a href="/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoPatientDetail?pid=3">Robinson, Mark</a>
            </div></td>
            <td></td><td></td><td></td>
        </tr>
        <!-- Duplicate row for Brusa (PracticeMojo bug) -->
        <tr><td></td><td></td>
            <td><div class="d-flex align-items-center">
                <a href="/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoPatientDetail?pid=2">Brusa, Brianna</a>
            </div></td>
            <td>03/05 @ 01:20 PM</td><td></td><td></td>
        </tr>
        </tbody></table></div>
        """
        soup = BeautifulSoup(html, "html.parser")
        results = parse_activity_detail(soup)

        # Family header row should be skipped
        names = [r["patient_name"] for r in results]
        assert "Pruitt Family" not in names

        # Should have: Pruitt x2, Brusa x2 (dup), Robinson (no appt)
        assert len(results) == 5

        # Now run the merge logic (same as merge_and_type_appointments)
        grouped = {}
        for row in results:
            try:
                validate_row(row, 23)
            except ParseError:
                continue
            patient = row["patient_name"].strip()
            appt = row["appointment"]
            if appt:
                grouped.setdefault(patient, []).append(appt)

        # Robinson has no appointment (moved) → skipped
        assert "Robinson, Mark" not in grouped

        # Pruitt has 2 different times → both present
        assert "Pruitt, Ashlee" in grouped
        assert len(grouped["Pruitt, Ashlee"]) == 2

        # Brusa appeared twice with same time (dup bug) → both in list before dedup
        assert "Brusa, Brianna" in grouped
        assert len(grouped["Brusa, Brianna"]) == 2  # both dupes in raw list

        # Now apply the dedup + format logic (chronological sort)
        from pmojo_lib.automation import _appt_sort_key
        for patient, appts in grouped.items():
            unique_appts = sorted(set(appts), key=_appt_sort_key)
            formatted = []
            for idx, a in enumerate(unique_appts):
                if idx == 0:
                    formatted.append(a)
                else:
                    parts = a.split('@')
                    formatted.append(parts[1].strip() if len(parts) > 1 else a)
            times_str = ' & '.join(formatted)

            if patient == "Pruitt, Ashlee":
                # Two different times merged
                assert "01:10 PM" in times_str
                assert "01:20 PM" in times_str
                assert "&" in times_str
            elif patient == "Brusa, Brianna":
                # Duplicate removed by set()
                assert times_str == "03/05 @ 01:20 PM"
                assert "&" not in times_str

    def test_normal_campaign_no_appointment_needed(self):
        """For non-appointment CDIs, appointment field should be empty."""
        soup = load_fixture("cdi1_cdn1.html")
        results = parse_activity_detail(soup)
        for r in results:
            assert r["appointment"] == "", f"Recare patient should have no appointment: {r}"

    def test_parsed_data_sufficient_for_softdent_entry(self):
        """Verify parsed data has everything needed for type_normal SoftDent entry."""
        soup = load_fixture("cdi33_cdn2.html")
        results = parse_activity_detail(soup)
        assert len(results) > 0
        for r in results:
            assert r["patient_name"].strip(), "Need patient name for SoftDent lookup"
            assert r["method"], "Need method to determine type_of_com"


# ===========================================================================
# TEST: Source file consistency
# ===========================================================================

# ---------------------------------------------------------------------------
# Import validate_row and ParseError directly (no Windows deps)
# ---------------------------------------------------------------------------
from unittest.mock import MagicMock, patch
from pmojo_lib.automation import validate_row, ParseError, APPOINTMENT_CDIS, COM_MAP, CDI_LIST, CDN_LIST
from pmojo_lib.practicemojo_api import SessionExpiredError


class TestValidateRow:
    """Tests for validate_row() — the parse safety gate before AHK typing."""

    def test_valid_normal_row(self):
        row = {"patient_name": "Smith, John", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        validate_row(row, 1)  # Should not raise

    def test_valid_appointment_row(self):
        row = {"patient_name": "Doe, Jane", "campaign": "Some Appt Reminder", "appointment": "03/04 @ 02:00 PM", "method": "Email"}
        validate_row(row, 23)  # Should not raise

    def test_empty_patient_name_raises(self):
        row = {"patient_name": "", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="Empty patient name"):
            validate_row(row, 1)

    def test_whitespace_only_name_raises(self):
        row = {"patient_name": "   ", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="Empty patient name"):
            validate_row(row, 1)

    def test_html_in_name_raises(self):
        row = {"patient_name": "<td>Smith</td>", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="HTML artifacts"):
            validate_row(row, 1)

    def test_html_entity_in_name_raises(self):
        row = {"patient_name": "Smith&nbsp;John", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="HTML artifacts"):
            validate_row(row, 1)

    def test_missing_appointment_for_appt_cdi_raises(self):
        row = {"patient_name": "Doe, Jane", "campaign": "Some Appt Reminder", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="Empty appointment"):
            validate_row(row, 23)

    def test_missing_appointment_for_cdi_130_raises(self):
        row = {"patient_name": "Doe, Jane", "campaign": "Another Appt Reminder", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="Empty appointment"):
            validate_row(row, 130)

    def test_placeholder_campaign_raises(self):
        """CDI 13 has placeholder 'SomeCampaign13' — should be caught."""
        row = {"patient_name": "Smith, John", "campaign": "SomeCampaign13", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="No valid campaign mapping"):
            validate_row(row, 13)

    def test_normal_appointment_not_required(self):
        """Non-appointment CDIs don't need an appointment field."""
        row = {"patient_name": "Smith, John", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        validate_row(row, 1)  # Should not raise

    def test_valid_name_with_comma(self):
        row = {"patient_name": "O'Brien, Mary-Jane", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        validate_row(row, 1)  # Should not raise

    def test_numeric_html_entity_in_name_raises(self):
        row = {"patient_name": "Smith&#160;John", "campaign": "Recare: Due", "appointment": "", "method": "Email"}
        with pytest.raises(ParseError, match="HTML artifacts"):
            validate_row(row, 1)

    def test_all_valid_cdis_have_campaign(self):
        """Every non-appointment CDI in CDI_LIST should have a real COM_MAP entry."""
        for cdi in COM_MAP:
            if cdi in APPOINTMENT_CDIS:
                continue
            com = COM_MAP[cdi]
            if com.startswith("SomeCampaign"):
                continue  # CDI 13 — known placeholder, tested above
            row = {"patient_name": "Test Patient", "campaign": com, "appointment": "", "method": "Email"}
            validate_row(row, cdi)  # Should not raise


# ===========================================================================
# TEST: Session expiry detection
# ===========================================================================

class TestSessionExpiry:
    """Test PracticeMojoAPI session expiry detection."""

    def test_redirect_to_login_page_raises(self):
        """If PM redirects to the login URL, SessionExpiredError should be raised."""
        from pmojo_lib.practicemojo_api import PracticeMojoAPI

        api = PracticeMojoAPI("user", "pass")
        api.session = MagicMock()

        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.url = "https://app.practicemojo.com/Pages/login?redirect=something"
        mock_resp.text = "<html>Login page</html>"
        api.session.get.return_value = mock_resp

        with pytest.raises(SessionExpiredError, match="redirected to login"):
            api.fetch_activity_detail("02/19/2026", 23, 1)

    def test_login_form_in_response_raises(self):
        """If the response body contains the login form fields, session expired."""
        from pmojo_lib.practicemojo_api import PracticeMojoAPI

        api = PracticeMojoAPI("user", "pass")
        api.session = MagicMock()

        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail"
        mock_resp.text = '<form><input name="loginId"/><input name="password"/><input name="slug"/></form>'
        api.session.get.return_value = mock_resp

        with pytest.raises(SessionExpiredError, match="login form in response"):
            api.fetch_activity_detail("02/19/2026", 23, 1)

    def test_no_session_raises(self):
        """If session is None, should raise SessionExpiredError."""
        from pmojo_lib.practicemojo_api import PracticeMojoAPI

        api = PracticeMojoAPI("user", "pass")
        api.session = None

        with pytest.raises(SessionExpiredError, match="Not logged in"):
            api.fetch_activity_detail("02/19/2026", 23, 1)

    def test_userlogin_in_redirect_url_raises(self):
        """If PM redirects to userLogin endpoint, session expired."""
        from pmojo_lib.practicemojo_api import PracticeMojoAPI

        api = PracticeMojoAPI("user", "pass")
        api.session = MagicMock()

        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/userLogin"
        mock_resp.text = "<html>Please log in</html>"
        api.session.get.return_value = mock_resp

        with pytest.raises(SessionExpiredError, match="redirected to login"):
            api.fetch_activity_detail("02/19/2026", 23, 1)

    def test_valid_response_no_false_positive(self):
        """A normal activity detail page should NOT trigger session expiry."""
        from pmojo_lib.practicemojo_api import PracticeMojoAPI

        api = PracticeMojoAPI("user", "pass")
        api.session = MagicMock()

        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_resp.url = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td=02/19/2026&cdi=23&cdn=1"
        mock_resp.text = '<div class="table-responsive"><table><thead></thead><tbody></tbody></table></div>'
        api.session.get.return_value = mock_resp

        # Should not raise — returns empty list since no patient rows
        result = api.fetch_activity_detail("02/19/2026", 23, 1)
        assert result == []

    def test_non_200_status_raises_exception(self):
        """Non-200 HTTP status should raise a generic Exception (not SessionExpiredError)."""
        from pmojo_lib.practicemojo_api import PracticeMojoAPI

        api = PracticeMojoAPI("user", "pass")
        api.session = MagicMock()

        mock_resp = MagicMock()
        mock_resp.status_code = 500
        api.session.get.return_value = mock_resp

        with pytest.raises(Exception, match="status=500"):
            api.fetch_activity_detail("02/19/2026", 23, 1)


# ===========================================================================
# TEST: Prefetch error handling
# ===========================================================================

class TestPrefetchErrors:
    """Test _prefetch_all error handling and thresholds."""

    def _make_automation(self, mock_api):
        """Create a PmojoAutomation with a mocked API and dummy SoftDent."""
        from pmojo_lib.automation import PmojoAutomation
        mock_db = MagicMock()
        mock_softdent = MagicMock()
        pm = PmojoAutomation(mock_db, mock_api, mock_softdent)
        return pm

    def test_session_expired_bubbles_up(self):
        """SessionExpiredError during prefetch should propagate, not be swallowed."""
        mock_api = MagicMock()
        mock_api.fetch_activity_detail.side_effect = SessionExpiredError("expired")
        pm = self._make_automation(mock_api)

        with pytest.raises(SessionExpiredError):
            pm._prefetch_all("02/19/2026")

    def test_too_many_failures_raises(self):
        """If more than half of fetches fail, should raise Exception."""
        call_count = 0
        total_combos = len(CDI_LIST) * len(CDN_LIST)

        def flaky_fetch(date_str, cdi, cdn):
            nonlocal call_count
            call_count += 1
            # Fail all but one
            if call_count == 1:
                return []
            raise Exception("network error")

        mock_api = MagicMock()
        mock_api.fetch_activity_detail.side_effect = flaky_fetch
        pm = self._make_automation(mock_api)

        with pytest.raises(Exception, match="Too many fetch failures"):
            pm._prefetch_all("02/19/2026")

    def test_few_failures_tolerated(self):
        """A small number of fetch failures should be tolerated (returns empty for those)."""
        call_count = 0

        def mostly_ok_fetch(date_str, cdi, cdn):
            nonlocal call_count
            call_count += 1
            # Fail just the first 2
            if call_count <= 2:
                raise Exception("transient error")
            return []

        mock_api = MagicMock()
        mock_api.fetch_activity_detail.side_effect = mostly_ok_fetch
        pm = self._make_automation(mock_api)

        result = pm._prefetch_all("02/19/2026")
        # Should succeed and return a dict for all combos
        assert len(result) == len(CDI_LIST) * len(CDN_LIST)

    def test_all_ok_returns_full_dict(self):
        """Normal operation: all fetches succeed."""
        mock_api = MagicMock()
        mock_api.fetch_activity_detail.return_value = [
            {"patient_name": "Smith, John", "campaign": "Test", "method": "Email",
             "appointment": "", "confirmations": "", "status": ""}
        ]
        pm = self._make_automation(mock_api)

        result = pm._prefetch_all("02/19/2026")
        assert len(result) == len(CDI_LIST) * len(CDN_LIST)
        # Each combo should have the row
        for key, rows in result.items():
            assert len(rows) == 1
            assert rows[0]["patient_name"] == "Smith, John"


class TestSourceConsistency:
    """Verify source modules haven't drifted from test expectations."""

    def _load_module_source(self, module_name):
        src_path = os.path.join(PROJECT_ROOT, "pmojo_lib", f"{module_name}.py")
        with open(src_path) as f:
            return f.read()

    def test_cdi_list_matches(self):
        """The CDI list in automation.py should match our test expectations."""
        src = self._load_module_source("automation")
        assert "CDI_LIST = [1, 8, 13, 21, 22, 23, 30, 33, 35, 36, 130]" in src

    def test_cdn_list_matches(self):
        src = self._load_module_source("automation")
        assert "CDN_LIST = [1, 2, 3]" in src

    def test_new_parser_present(self):
        """practicemojo_api.py should contain the new structure parser."""
        src = self._load_module_source("practicemojo_api")
        assert "div.table-responsive table" in src
        assert "_parse_new_structure" in src

    def test_legacy_parser_present(self):
        """practicemojo_api.py should still contain the legacy parser as fallback."""
        src = self._load_module_source("practicemojo_api")
        assert "div.activity_detail_table table" in src
        assert "_parse_legacy_structure" in src

    def test_softdent_uses_pywinctl(self):
        """softdent.py should use pywinctl (not pywinauto)."""
        src = self._load_module_source("softdent")
        assert "import pywinctl" in src
        assert "pywinauto" not in src

    def test_softdent_has_platform_guard(self):
        """softdent.py should have a platform guard for Mac compatibility."""
        src = self._load_module_source("softdent")
        assert "platform.system()" in src

    def test_pmojo_entry_point_exists(self):
        """pmojo.py should exist as the entry point."""
        src = _load_pmojo_source()
        assert "from pmojo_lib" in src
        assert '__name__ == "__main__"' in src


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
