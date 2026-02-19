"""Core automation: fetches data from PracticeMojo and types it into SoftDent."""
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime as _dt
from time import sleep

from pmojo_lib.events import STOP_EVENT, PAUSE_EVENT
from pmojo_lib.database import Database
from pmojo_lib.practicemojo_api import PracticeMojoAPI, SessionExpiredError
from pmojo_lib.softdent import SoftDentConnection


class ParseError(Exception):
    """Raised when parsed data fails validation — stops all AHK typing immediately."""
    pass

# Campaign name mapping for SoftDent entry
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

CDI_LIST = [1, 8, 13, 21, 22, 23, 30, 33, 35, 36, 130]
CDN_LIST = [1, 2, 3]

# CDIs that are appointment reminders (use merge logic)
APPOINTMENT_CDIS = [23, 130]

# CDIs that get the recare "r" flag
RECARE_CDIS = [1, 8, 33]


_HTML_ARTIFACTS = re.compile(r'<[^>]+>|&[a-zA-Z]+;|&\#\d+;')


def _appt_sort_key(appt_str):
    """Sort key for appointment strings like '03/05 @ 01:10 PM' — chronological order."""
    try:
        parts = appt_str.split('@')
        time_part = parts[1].strip() if len(parts) > 1 else appt_str.strip()
        return _dt.strptime(time_part, "%I:%M %p")
    except (ValueError, IndexError):
        # Fallback: sort to end if format is unexpected
        return _dt.max


def validate_row(row, cdi):
    """Validate a parsed row before AHK types it. Raises ParseError if invalid.

    Checks:
    - patient_name is non-empty after stripping
    - patient_name doesn't contain HTML artifacts (indicates broken parse)
    - Appointment CDIs must have a non-empty appointment field
    - COM_MAP must have a real campaign name for non-appointment CDIs
    """
    name = row.get("patient_name", "").strip()
    if not name:
        raise ParseError(
            f"Empty patient name in parsed row (cdi={cdi}). "
            f"Row data: {row}"
        )

    if _HTML_ARTIFACTS.search(name):
        raise ParseError(
            f"HTML artifacts in patient name '{name}' (cdi={cdi}). "
            f"PracticeMojo HTML structure may have changed."
        )

    if cdi in APPOINTMENT_CDIS:
        appt = row.get("appointment", "").strip()
        if not appt:
            raise ParseError(
                f"Empty appointment for patient '{name}' (cdi={cdi}). "
                f"Row data: {row}"
            )

    com = COM_MAP.get(cdi, "")
    if not com or com.startswith("SomeCampaign"):
        if cdi not in APPOINTMENT_CDIS:
            raise ParseError(
                f"No valid campaign mapping for cdi={cdi} (got '{com}'). "
                f"Patient '{name}' would receive incorrect data."
            )


def determine_type_of_com(cdi, cdn):
    """Map CDI/CDN to communication type code for SoftDent entry."""
    if cdi in APPOINTMENT_CDIS:
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


class PmojoAutomation:
    """Orchestrates fetching data from PM and typing it into SoftDent for each date/cdi/cdn."""

    def __init__(self, db: Database, pm_api: PracticeMojoAPI, softdent: SoftDentConnection):
        self.db = db
        self.pm_api = pm_api
        self.softdent = softdent

    def _prefetch_all(self, date_str: str):
        """Fetch all CDI/CDN pages concurrently. Returns dict of (cdi, cdn) -> rows.

        Raises SessionExpiredError immediately if any fetch detects an expired session.
        Other fetch errors are tracked; raises Exception if too many fail.
        """
        prefetched = {}
        combos = [(cdi, cdn) for cdi in CDI_LIST for cdn in CDN_LIST]
        fetch_errors = 0

        def fetch_one(cdi, cdn):
            return (cdi, cdn, self.pm_api.fetch_activity_detail(date_str, cdi, cdn))

        with ThreadPoolExecutor(max_workers=8) as executor:
            futures = {executor.submit(fetch_one, cdi, cdn): (cdi, cdn) for cdi, cdn in combos}
            for future in as_completed(futures):
                cdi, cdn = futures[future]
                try:
                    _, _, rows = future.result()
                    prefetched[(cdi, cdn)] = rows
                except SessionExpiredError:
                    # Cancel remaining futures and bubble up immediately
                    for f in futures:
                        f.cancel()
                    raise
                except Exception as e:
                    fetch_errors += 1
                    print(f"[{date_str}] Error fetching cdi={cdi}, cdn={cdn}: {e}")
                    prefetched[(cdi, cdn)] = []

        if fetch_errors > 0:
            print(f"[{date_str}] WARNING: {fetch_errors}/{len(combos)} fetches failed.")
            # If most fetches failed, something is wrong (network issue, etc.)
            if fetch_errors > len(combos) // 2:
                raise Exception(
                    f"Too many fetch failures ({fetch_errors}/{len(combos)}) for {date_str}. "
                    f"Possible network issue."
                )

        return prefetched

    def process_date(self, date_str: str):
        """Main method: prefetch all pages concurrently, then type sequentially."""
        m, d, y = date_str.split('/')

        # Fetch all 33 CDI/CDN pages in parallel before touching SoftDent
        print(f"[{date_str}] Prefetching all campaign data...")
        prefetched = self._prefetch_all(date_str)
        print(f"[{date_str}] Prefetch complete.")

        # Validate all prefetched data before connecting to SoftDent
        for cdi in CDI_LIST:
            for cdn in CDN_LIST:
                rows = prefetched.get((cdi, cdn), [])
                if not rows:
                    continue

                # Structural validation: if ALL rows fail, the HTML format likely changed
                all_invalid = True
                first_error = None
                for row in rows:
                    try:
                        validate_row(row, cdi)
                        all_invalid = False
                        break
                    except ParseError as e:
                        if first_error is None:
                            first_error = e
                if all_invalid and first_error is not None:
                    STOP_EVENT.set()
                    self.db.toggle_day_status(date_str, 'error')
                    raise ParseError(
                        f"ALL {len(rows)} rows failed validation for cdi={cdi}, cdn={cdn} on {date_str}. "
                        f"First error: {first_error}"
                    )

        self.softdent.connect()
        self.softdent.focus()

        for cdi in CDI_LIST:
            for cdn in CDN_LIST:
                if STOP_EVENT.is_set():
                    self.db.toggle_day_status(date_str, 'error')
                    print(f"Stop pressed during {date_str}, marking error.")
                    return False
                PAUSE_EVENT.wait()

                rows = prefetched.get((cdi, cdn), [])

                if not rows:
                    print(f"[{date_str}] No data for cdi={cdi}, cdn={cdn}.")
                else:
                    print(f"[{date_str}] Found {len(rows)} rows for cdi={cdi}, cdn={cdn}.")
                    print(f"[{date_str}] Campaign: {rows[0]['campaign']}, Method: {rows[0]['method']}")
                    print(f"[{date_str}] Patients: {[row['patient_name'] for row in rows]}")

                if cdi in APPOINTMENT_CDIS:
                    self.merge_and_type_appointments(rows, cdi, cdn, m, d, y, date_str)
                else:
                    self.type_normal(rows, cdi, cdn, m, d, y, date_str)

        if not STOP_EVENT.is_set():
            self.db.toggle_day_status(date_str, 'done')
            print(f"[{date_str}] Marked as 'done'.")
            return True
        else:
            self.db.toggle_day_status(date_str, 'error')
            print(f"[{date_str}] Marked as 'error' after partial processing.")
            return False

    def merge_and_type_appointments(self, rows, cdi, cdn, m, d, y, date_str):
        """Group times by patient, combine times, type into SoftDent."""
        grouped = {}
        for row in rows:
            try:
                validate_row(row, cdi)
            except ParseError as e:
                print(f"[{date_str}] Skipping invalid row: {e}")
                continue
            patient = row["patient_name"].strip()
            appt = row["appointment"]
            if appt:
                grouped.setdefault(patient, []).append(appt)

        type_of_com = determine_type_of_com(cdi, cdn)

        for patient, appts in grouped.items():
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return
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
            times_str = ' & '.join(formatted)

            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type("f")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type(patient)
            sleep(2)
            self.softdent.ahk.key_press("Enter")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type("0ca")
            self.softdent.ahk.type("cc")
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type("Reminder for")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type(" " + times_str)
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type(type_of_com)
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type("pmojoNFD")
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type(f"{m}/{d}/{y}")
            sleep(3)
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.key_press("Enter")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type("l")
            self.softdent.ahk.type("l")
            self.softdent.ahk.type("n")

    def type_normal(self, rows, cdi, cdn, m, d, y, date_str):
        """Type normal campaign data into SoftDent."""
        com = COM_MAP.get(cdi, "")
        type_of_com = determine_type_of_com(cdi, cdn)
        recare_flag = "r" if cdi in RECARE_CDIS else ""

        for row in rows:
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return
            try:
                validate_row(row, cdi)
            except ParseError as e:
                print(f"[{date_str}] Skipping invalid row: {e}")
                continue
            patient = row["patient_name"].strip()

            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type("f")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type(patient)
            sleep(2)
            self.softdent.ahk.key_press("Enter")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type("0ca")
            self.softdent.ahk.type(recare_flag)
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type(com)
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type(type_of_com)
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type("pmojoNFD")
            self.softdent.ahk.key_press("Tab")
            self.softdent.ahk.type(f"{m}/{d}/{y}")
            sleep(3)
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.key_press("Enter")
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return

            self.softdent.ahk.type("l")
            self.softdent.ahk.type("l")
            self.softdent.ahk.type("n")
