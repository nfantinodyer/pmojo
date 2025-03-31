import tkinter as tk
import re
import json
import warnings
import threading
import keyboard
import calendar
import sqlite3
import datetime
import requests
from bs4 import BeautifulSoup
from time import sleep
from pywinauto.application import Application
from ahk import AHK

warnings.simplefilter("ignore", category=UserWarning)

STOP_EVENT = threading.Event()
PAUSE_EVENT = threading.Event()
PAUSE_EVENT.set()

class Database:
    def __init__(self, db_path="days.db", gui=None):
        self.db_path = db_path
        self.gui = gui
        self.init_db()

    @staticmethod
    def _ui_str_to_date(ui_str: str) -> datetime.date:
        """
        Converts 'MM/DD/YYYY' (UI format) -> datetime.date.
        """
        return datetime.datetime.strptime(ui_str, "%m/%d/%Y").date()

    @staticmethod
    def _date_to_db_str(d: datetime.date) -> str:
        """
        Converts datetime.date -> 'YYYY-MM-DD' for storing in DB.
        """
        return d.strftime("%Y-%m-%d")

    @staticmethod
    def _db_str_to_date(db_str: str) -> datetime.date:
        """
        Converts 'YYYY-MM-DD' (DB format) -> datetime.date.
        """
        return datetime.datetime.strptime(db_str, "%Y-%m-%d").date()

    def init_db(self):
        """
        Ensures the 'days' table exists and pre-populates from 2020-01-01 to 2025-03-27 as 'done',
        storing each date in 'YYYY-MM-DD' format.
        """
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
        end_date   = datetime.date(2025, 3, 27)
        delta = datetime.timedelta(days=1)

        to_insert = []
        current = start_date
        while current <= end_date:
            iso_str = self._date_to_db_str(current)
            to_insert.append((iso_str, "done", ""))
            current += delta

        c.executemany("""
            INSERT OR IGNORE INTO days(date, status, error_msg)
            VALUES (?, ?, ?)
        """, to_insert)

        conn.commit()
        conn.close()

    def expand_db_up_to(self, end_date: datetime.date):
        """
        Inserts rows in 'days' table up to 'end_date' if not already present.
        Uses 'YYYY-MM-DD' format so MAX(date) works correctly.
        """
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
                c.executemany("""
                    INSERT OR IGNORE INTO days(date, status, error_msg)
                    VALUES (?, ?, ?)
                """, to_insert)
        conn.commit()
        conn.close()

    def get_day_status(self, ui_date_str: str) -> str:
        """
        Return the 'status' field for ui_date_str='MM/DD/YYYY' from the DB,
        or '' if not found.
        """
        try:
            d = self._ui_str_to_date(ui_date_str)      # parse UI date
            iso_str = self._date_to_db_str(d)          # convert to ISO
        except ValueError:
            return ""

        conn = sqlite3.connect(self.db_path)
        c = conn.cursor()
        c.execute("SELECT status FROM days WHERE date=?", (iso_str,))
        row = c.fetchone()
        conn.close()

        return row[0] if row else ""

    def toggle_day_status(self, ui_date_str: str, new_status: str):
        """
        Toggle the status of the given ui_date_str='MM/DD/YYYY' to new_status.
        If the status is already new_status, do nothing.
        """
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

        if self.gui is not None:
            self.gui.safe_update_calendar()

class PracticeMojoAPI:
    """
    Handles PracticeMojo login and fetching data.
    """
    BASE_URL = "https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/"

    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.session = None

    def login(self):
        """
        Logs in and sets self.session to an authenticated requests.Session.
        """
        s = requests.Session()
        login_page_url = "https://app.practicemojo.com/Pages/login"
        s.get(login_page_url)

        # Note: The base login endpoint may differ depending on your site structure.
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
        """
        Fetch the detail page for (date_str, cdi, cdn) from PM, parse and return the row dicts.
        """
        if not self.session:
            raise Exception("Not logged in to PracticeMojo.")
        detail_url = f"{self.BASE_URL}gotoActivityDetail?td={date_str}&cdi={cdi}&cdn={cdn}"
        resp = self.session.get(detail_url)
        if resp.status_code != 200:
            raise Exception(f"Failed to fetch detail page (cdi={cdi}, cdn={cdn}): status={resp.status_code}")
        soup = BeautifulSoup(resp.text, "html.parser")
        return self.parse_activity_detail(soup)

    @staticmethod
    def parse_activity_detail(soup: BeautifulSoup):
        """
        Parse the HTML soup for the campaign detail table and return a list of dicts.
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
                        patient_name  = data_tds[0].get_text().replace("↳", "").rstrip()
                        if "Family" not in patient_name:
                            appointment   = data_tds[1].get_text(strip=True)
                            confirmations = data_tds[2].get_text(strip=True)
                            row_status    = data_tds[3].get_text(strip=True)
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

class PywinAuto:
    """
    Connect to SoftDent (or fallback WordPad) and handle AHK typing.
    """
    def __init__(self):
        self.ahk = AHK()
        self.app = Application()
        self.win = None  # will store the connected window

    def connect(self):
        """
        Connect to SoftDent, else fallback to WordPad, store window in self.win.
        """
        try:
            self.app.connect(title_re=".*CS SoftDent.*- S.*")
            self.win = self.app.window(title_re=".*CS SoftDent.*- S.*")
            print("Connected to SoftDent.")
        except Exception:
            print("SoftDent not found. Connecting to WordPad for testing.")
            self.app.connect(title_re=".*WordPad.*")
            self.win = self.app.window(title_re=".*WordPad.*")

    def focus(self):
        if self.win is not None:
            self.win.set_focus()

    def type_text(self, text: str):
        self.ahk.type(text)

    def press_tab(self, times=1):
        for _ in range(times):
            self.ahk.key_press("Tab")

    # You can add more convenience methods as needed, e.g. press_enter, etc.

class PmojoAutomation:
    """
    Orchestrates fetching data from PM and typing it into SoftDent for each date/cdi/cdn.
    """
    def __init__(self, db: Database, pm_api: PracticeMojoAPI, softdent: PywinAuto):
        self.db = db
        self.pm_api = pm_api
        self.softdent = softdent

    def process_date(self, date_str: str):
        """
        Main method: for each cdi/cdn, fetch detail, type it in.
        """
        m, d, y = date_str.split('/')
        cdi_list = [1, 8, 13, 21, 22, 23, 30, 33, 35, 36, 130]
        cdn_list = [1, 2, 3]

        # Make sure we have a connected SoftDent/WordPad window
        self.softdent.connect()
        self.softdent.focus()

        for cdi in cdi_list:
            for cdn in cdn_list:
                if STOP_EVENT.is_set():
                    # user pressed 'Stop'
                    self.db.toggle_day_status(date_str, 'error')
                    print(f"Stop pressed during {date_str}, marking error.")
                    return False
                PAUSE_EVENT.wait()

                # fetch data
                try:
                    rows = self.pm_api.fetch_activity_detail(date_str, cdi, cdn)
                except Exception as e:
                    print(f"[{date_str}] Error fetching cdi={cdi}, cdn={cdn}: {e}")
                    continue

                if not rows:
                    print(f"[{date_str}] No data for cdi={cdi}, cdn={cdn}.")
                else:
                    print(f"[{date_str}] Found {len(rows)} rows for cdi={cdi}, cdn={cdn}.")
                    print(f"[{date_str}] Campaign: {rows[0]['campaign']}, Method: {rows[0]['method']}")
                    print(f"[{date_str}] Patients: {[row['patient_name'] for row in rows]}")

                # route to merge or normal
                if cdi in [23, 130]:
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
        """
        nameDateTime logic: group times by patient, combine times, type 0ca + cc + ...
        """
        grouped = {}
        for row in rows:
            patient = row["patient_name"].strip()
            appt = row["appointment"]
            if appt:
                grouped.setdefault(patient, []).append(appt)

        type_of_com = self.determine_type_of_com(cdi, cdn)

        for patient, appts in grouped.items():
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return
            unique_appts = sorted(set(appts))
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

            # AHK typing
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
        """
        name(...) logic for normal campaigns (0ca + recare_flag + com => typeOfCom => pmojoNFD => date).
        """
        com_map = {
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
        com = com_map.get(cdi, "")
        type_of_com = self.determine_type_of_com(cdi, cdn)
        recare_flag = "r" if cdi in [1, 8, 33] else ""

        for row in rows:
            PAUSE_EVENT.wait()
            if STOP_EVENT.is_set():
                return
            patient = row["patient_name"].strip()

            # AHK typing
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
    
    def determine_type_of_com(self, cdi, cdn):
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

class PmojoGUI:
    def __init__(self, db: Database):
        self.db = db
        self.root = tk.Tk()
        self.root.withdraw()
        self.window = tk.Toplevel(self.root)
        self.window.title("Pmojo")
        self.completeAllVar = tk.BooleanVar(value=False)
        self.entry = None
        self.cur_month = 0
        self.cur_year = 0
        self.create_gui()

    def run(self):
        self.window.mainloop()

    def create_gui(self):
        self.is_running = False

        #references to widgets
        self.btn_done = None
        self.btn_error = None
        self.btn_left = None
        self.btn_right = None
        self.chk_complete_all = None
        self.btn_start = None
        self.btn_exit = None

        self.window.columnconfigure(0, weight=1, minsize=250)
        self.window.rowconfigure(0, weight=1, minsize=250)

        label = tk.Label(self.window, text="Enter the date (MM/DD/YYYY): ")
        label.pack(side=tk.TOP, anchor=tk.W)

        frame1 = tk.Frame(self.window, relief=tk.SUNKEN)
        frame1.pack(fill=tk.X)

        # Default to today
        default_date = datetime.date.today().strftime("%m/%d/%Y")
        self.entry = tk.Entry(frame1, width=50)
        self.entry.pack(side=tk.LEFT, padx=2, pady=5)
        self.entry.insert(0, default_date)
        self.entry.focus_set()
        self.entry.bind("<Return>", lambda e: self.begin(self.entry.get()))

        # completeAll checkbox
        self.chk_complete_all = tk.Checkbutton(
            self.window,
            text="Complete All",
            variable=self.completeAllVar,
            command=self.update_calendar
        )
        self.chk_complete_all.pack(anchor=tk.W)

        # Parse month/year from default
        m, d, y = default_date.split('/')
        self.cur_month = int(m)
        self.cur_year = int(y)

        # Month nav
        month_frame = tk.Frame(self.window)
        month_frame.pack(pady=(5, 0), anchor="center")

        self.btn_left = tk.Button(month_frame, text="←", command=self.prev_month)
        self.btn_left.pack(side=tk.LEFT, padx=5)

        self.month_label = tk.Label(month_frame, text="", font=("Helvetica", 12, "bold"))
        self.month_label.pack(side=tk.LEFT)

        self.btn_right = tk.Button(month_frame, text="→", command=self.next_month)
        self.btn_right.pack(side=tk.LEFT, padx=5)

        self.days_frame = tk.Frame(self.window)
        self.days_frame.pack(side=tk.TOP, anchor="center", padx=2, pady=5)

        # status area
        status_frame = tk.Frame(self.window)
        status_frame.pack(anchor='center', pady=5)

        lbl_status = tk.Label(status_frame, text="Status: ")
        lbl_status.pack(side=tk.LEFT, padx=5)

        self.btn_done = tk.Button(status_frame, text="Toggle Done", command=self.mark_done)
        self.btn_done.pack(side=tk.LEFT, padx=5)

        self.btn_error = tk.Button(status_frame, text="Toggle Error", command=self.mark_error)
        self.btn_error.pack(side=tk.LEFT, padx=5)

        # bind event
        self.entry.bind("<KeyRelease>", self.update_calendar)
        self.draw_calendar(self.cur_month, self.cur_year)

        # Bottom buttons
        btns_frame = tk.Frame(self.window)
        btns_frame.pack(side=tk.BOTTOM, pady=10)

        self.btn_start = tk.Button(btns_frame, text="Start", width=10, height=2, command=self.start_button_action)
        self.btn_start.pack(side=tk.LEFT, padx=5)

        self.btn_exit = tk.Button(btns_frame, text="Exit", width=10, height=2, command=self.on_exit)
        self.btn_exit.pack(side=tk.LEFT, padx=5)

        # Global key bindings
        self.window.bind_all("<Tab>", lambda e: self.window.tk_focusNext().focus())
        self.window.bind_all("<Return>", lambda e: self.window.focus_get().invoke())

        # finalize
        self.update_calendar()

    def set_running(self, running: bool):
        """
        If running is True, disable all buttons except 'Exit'.
        If running is False, enable them again.
        """
        self.is_running = running
        
        new_state = "disabled" if running else "normal"
        
        # disable/enable the Start button
        if self.btn_start:
            self.btn_start.config(state=new_state)

        # disable/enable Toggle Done / Error
        if self.btn_done:
            self.btn_done.config(state=new_state)
        if self.btn_error:
            self.btn_error.config(state=new_state)

        # disable/enable month nav
        if self.btn_left:
            self.btn_left.config(state=new_state)
        if self.btn_right:
            self.btn_right.config(state=new_state)

        # disable/enable the "Complete All" checkbox
        if self.chk_complete_all:
            if running:
                self.chk_complete_all.config(state="disabled")
            else:
                # For a checkbutton, 'normal' is the correct enable state
                self.chk_complete_all.config(state="normal")

        # disable/enable day buttons for the current month
        if hasattr(self, "day_buttons"):
            for btn in self.day_buttons:
                btn.config(state=new_state)

    def prev_month(self):
        self.cur_month -= 1
        if self.cur_month < 1:
            self.cur_month = 12
            self.cur_year -= 1
        self.draw_calendar(self.cur_month, self.cur_year)

    def next_month(self):
        self.cur_month += 1
        if self.cur_month > 12:
            self.cur_month = 1
            self.cur_year += 1
        self.draw_calendar(self.cur_month, self.cur_year)

    def draw_calendar(self, month, year):
        # Clear previous day buttons
        for widget in self.days_frame.winfo_children():
            widget.destroy()

        self.month_label.config(text=f"{calendar.month_name[month]} {year}")

        # Possibly expand DB if "Complete All" is checked
        text = self.entry.get().strip()
        green_dates = set()

        # If entry is a valid MM/DD/YYYY, add it initially
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
            green_dates.add(text)

        if self.completeAllVar.get():
            try:
                # parse the end date from the entry
                mm, dd, yy = text.split('/')
                end_date = datetime.date(int(yy), int(mm), int(dd))

                # Expand database if needed
                self.db.expand_db_up_to(end_date)

                # Gather all not-done dates
                conn = sqlite3.connect(self.db.db_path)
                c = conn.cursor()
                c.execute("SELECT date FROM days WHERE status != 'done'")
                rows = c.fetchall()
                conn.close()

                # Convert ISO -> datetime -> UI "MM/DD/YYYY"
                for (db_date_str,) in rows:
                    try:
                        dt = self.db._db_str_to_date(db_date_str)
                        if dt <= end_date:
                            ui_str = dt.strftime("%m/%d/%Y")
                            green_dates.add(ui_str)
                    except:
                        pass
            except:
                pass

        # draw day-of-week labels
        days_of_the_week = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        for col_idx, dow in enumerate(days_of_the_week):
            lbl = tk.Label(self.days_frame, text=dow, width=3, font=("Helvetica", 10, "bold"))
            lbl.grid(row=0, column=col_idx, padx=2, pady=2)

        cal_iter = calendar.Calendar(firstweekday=calendar.SUNDAY)
        row_idx = 1
        for week in cal_iter.monthdayscalendar(year, month):
            col_idx = 0
            for day in week:
                if day == 0:
                    # Empty cell
                    tk.Label(self.days_frame, text=" ", width=3).grid(row=row_idx, column=col_idx)
                else:
                    ds = f"{month:02}/{day:02}/{year}"
                    status = self.db.get_day_status(ds)
                    btn_kwargs = {"text": str(day), "width": 3}

                    if status == "done":
                        btn_kwargs["bg"] = "gray"
                    elif status == "error" and ds in green_dates:
                        btn_kwargs["bg"] = "green"
                        btn_kwargs["highlightbackground"] = "red"
                    elif status == "error":
                        btn_kwargs["bg"] = "red"
                    else:
                        if ds in green_dates:
                            btn_kwargs["bg"] = "lightgreen"
                        else:
                            btn_kwargs["bg"] = "SystemButtonFace"

                    def on_day_click(ds=ds):
                        self.entry.delete(0, tk.END)
                        self.entry.insert(0, ds)
                        self.update_calendar()

                    tk.Button(self.days_frame, command=on_day_click, **btn_kwargs).grid(row=row_idx, column=col_idx)
                col_idx += 1
            row_idx += 1

    def update_calendar(self, event=None):
        text = self.entry.get().strip()
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
            try:
                new_m, new_d, new_y = text.split('/')
                new_m = int(new_m)
                new_y = int(new_y)
                if 1 <= new_m <= 12:
                    self.cur_month, self.cur_year = new_m, new_y
                    self.draw_calendar(new_m, new_y)
                else:
                    self.draw_calendar(self.cur_month, self.cur_year)
            except:
                self.draw_calendar(self.cur_month, self.cur_year)
        else:
            self.draw_calendar(self.cur_month, self.cur_year)

    def safe_update_calendar(self):
        if self.root and self.root.winfo_exists() and self.update_calendar is not None:
            self.root.after(1, self.update_calendar)
    
    def mark_done(self):
        date_str = self.entry.get().strip()
        old = self.db.get_day_status(date_str)
        new = "" if old == "done" else "done"
        self.db.toggle_day_status(date_str, new)
        self.update_calendar()

    def mark_error(self):
        date_str = self.entry.get().strip()
        old = self.db.get_day_status(date_str)
        new = "" if old == "error" else "error"
        self.db.toggle_day_status(date_str, new)
        self.update_calendar()

    def start_button_action(self):
        self.set_running(True)

        if self.completeAllVar.get():
            threading.Thread(target=self.complete_range).start()
        else:
            self.begin(self.entry.get())

    def on_exit(self):
        STOP_EVENT.set()
        keyboard.unhook_all_hotkeys()
        self.window.quit()

    def begin(self, date_str: str):
        STOP_EVENT.clear()
        def background():
            try:
                with open("config.json") as cfg:
                    conf = json.load(cfg)
                pm_api = PracticeMojoAPI(conf["USERNAME"], conf["PASSWORD"])
                pm_api.login()
            except Exception as e:
                print(f"PracticeMojo login failed: {e}")
                self.db.toggle_day_status(date_str, "error")
                return

            automator = PywinAuto()
            pm = PmojoAutomation(self.db, pm_api, automator)
            pm.process_date(date_str)
            self.window.after(0, lambda: self.set_running(False))

        threading.Thread(target=background).start()

    def complete_range(self):
        # read end date from entry
        end_str = self.entry.get().strip()
        try:
            mm, dd, yy = end_str.split('/')
            end_date = datetime.date(int(yy), int(mm), int(dd))
        except:
            print("Invalid date in Entry, cannot complete range.")
            return

        # gather undone dates (which are stored as ISO 'YYYY-MM-DD')
        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        c.execute("SELECT date FROM days WHERE status != 'done'")
        rows = c.fetchall()
        conn.close()

        undone = []
        # Convert each row from ISO -> datetime -> UI for logging
        for (db_date_str,) in rows:
            try:
                dt = self.db._db_str_to_date(db_date_str)
                if dt <= end_date:
                    # Store as (datetime_object, "MM/DD/YYYY")
                    ds_ui = dt.strftime("%m/%d/%Y")
                    undone.append((dt, ds_ui))
            except:
                pass

        undone.sort(key=lambda x: x[0])
        print(f"Completing {len(undone)} dates up to {end_str}...")

        def background():
            STOP_EVENT.clear()
            # login
            try:
                with open("config.json") as cfg:
                    conf = json.load(cfg)
                pm_api = PracticeMojoAPI(conf["USERNAME"], conf["PASSWORD"])
                pm_api.login()
            except Exception as e:
                print(f"PracticeMojo login failed: {e}")
                return

            automator = PywinAuto()
            pm = PmojoAutomation(self.db, pm_api, automator)

            for (dt_obj, ds_ui) in undone:
                if STOP_EVENT.is_set():
                    self.db.toggle_day_status(ds_ui, "error")
                    print(f"Stop pressed before processing {ds_ui}, marking error.")
                    return
                print(f"Processing date {ds_ui}...")
                try:
                    success = pm.process_date(ds_ui)
                    if not success or STOP_EVENT.is_set():
                        self.db.toggle_day_status(ds_ui, "error")
                        print(f"Marked {ds_ui} as 'error' (canceled).")
                        break
                    else:
                        self.db.toggle_day_status(ds_ui, "done")
                        print(f"Marked {ds_ui} as 'done'.")
                except Exception as ex:
                    self.db.toggle_day_status(ds_ui, "error")
                    print(f"Error on {ds_ui}: {ex}")
            STOP_EVENT.clear()
            self.window.after(0, lambda: self.set_running(False))
            print("Complete range finished.")

        threading.Thread(target=background).start()

def stop_program():
    STOP_EVENT.set()
    print("Stop triggered")

def pause_program():
    if PAUSE_EVENT.is_set():
        PAUSE_EVENT.clear()
        print("Paused (scroll lock)")
    else:
        PAUSE_EVENT.set()
        print("Resumed (scroll lock)")

keyboard.add_hotkey("num lock", stop_program)
keyboard.add_hotkey("scroll lock", pause_program)

if __name__ == "__main__":
    db = Database("days.db")
    gui = PmojoGUI(db)
    db.gui = gui
    gui.run()