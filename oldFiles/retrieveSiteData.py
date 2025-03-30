from http.server import executable
# Removed selenium imports
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import tkinter as tk
import re
from ahk import AHK
from time import sleep
import json
import warnings
import threading
import keyboard
import calendar
import sqlite3
import datetime
import sys

import requests
from bs4 import BeautifulSoup

warnings.simplefilter('ignore', category=UserWarning)  # Ignore 32bit warnings

ahk = AHK()
pauseEvent = threading.Event()
pauseEvent.set()
stopNow = threading.Event()
root = tk.Tk()
root.withdraw()
MAIN_UPDATE_CALENDAR = None

def stopProgram():
    stopNow.set()

def get_clipboard():
    return root.clipboard_get()

def expand_db_up_to(end_date):
    """
    Ensure the 'days' table has rows for every date from the last known entry
    up to (and including) 'end_date', if 'end_date' is beyond what's currently in the DB.
    We'll insert them with status = '' (undone).
    """
    conn = sqlite3.connect("days.db")
    c = conn.cursor()

    c.execute("SELECT MAX(date) FROM days")
    row = c.fetchone()
    latest_str = row[0] if row and row[0] else None

    if latest_str:
        m, d, y = latest_str.split('/')
        latest_date = datetime.date(int(y), int(m), int(d))
    else:
        latest_date = datetime.date(2020, 1, 1)

    if end_date > latest_date:
        one_day = datetime.timedelta(days=1)
        cur_date = latest_date + one_day
        while cur_date <= end_date:
            date_str = cur_date.strftime("%m/%d/%Y")
            c.execute("""
                INSERT OR IGNORE INTO days(date, status, error_msg)
                VALUES (?, '', '')
            """, (date_str,))
            cur_date += one_day

    conn.commit()
    conn.close()

def createGUI():
    window = tk.Toplevel(root)
    window.title("Pmojo")
    window.columnconfigure(0, weight=1, minsize=250)
    window.rowconfigure(0, weight=1, minsize=250)

    entrylabel = tk.Label(window, text="Enter the date (MM/DD/YYYY): ")
    entrylabel.pack(side=tk.TOP, anchor=tk.W)

    frame1 = tk.Frame(window, width=50, relief=tk.SUNKEN)
    frame1.pack(fill=tk.X)

    entry = tk.Entry(frame1, width=50)
    entry.pack(side=tk.LEFT, padx=2, pady=5)
    entry.insert(0, calendar.datetime.date.today().strftime("%m/%d/%Y"))
    entry.focus_set()
    entry.bind("<Return>", lambda e: begin(entry.get()))

    completeAllVar = tk.BooleanVar(value=False)
    def on_complete_all_toggle():
        update_calendar()

    completeAllCheck = tk.Checkbutton(
        window,
        text="Complete All",
        variable=completeAllVar,
        command=on_complete_all_toggle
    )
    completeAllCheck.pack(anchor=tk.W)

    cur_date_str = entry.get()
    m, d, y = cur_date_str.split('/')
    cur_month = int(m)
    cur_year = int(y)

    month_frame = tk.Frame(window)
    month_frame.pack(pady=(5, 0), anchor="center")

    def prev_month():
        nonlocal cur_month, cur_year
        cur_month -= 1
        if cur_month < 1:
            cur_month = 12
            cur_year -= 1
        draw_calendar(cur_month, cur_year)

    def next_month():
        nonlocal cur_month, cur_year
        cur_month += 1
        if cur_month > 12:
            cur_month = 1
            cur_year += 1
        draw_calendar(cur_month, cur_year)

    left_arrow_btn = tk.Button(month_frame, text="←", command=prev_month)
    left_arrow_btn.pack(side=tk.LEFT, padx=5)

    month_label = tk.Label(month_frame, text="", font=("Helvetica", 12, "bold"))
    month_label.pack(side=tk.LEFT)

    right_arrow_btn = tk.Button(month_frame, text="→", command=next_month)
    right_arrow_btn.pack(side=tk.LEFT, padx=5)

    days_frame = tk.Frame(window)
    days_frame.pack(side=tk.TOP, anchor="center", padx=2, pady=5)

    status_frame = tk.Frame(window)
    status_frame.pack(anchor='center', pady=5)

    selected_day_label = tk.Label(status_frame, text="Status: ")
    selected_day_label.pack(side=tk.LEFT, padx=5)

    def mark_done():
        date_str = entry.get().strip()
        if get_day_status(date_str) == 'done':
            toggle_day_status(date_str, '')
            update_calendar()
        else:
            toggle_day_status(date_str, 'done')
            update_calendar()

    def mark_error():
        date_str = entry.get().strip()
        if get_day_status(date_str) == 'error':
            toggle_day_status(date_str, '')
            update_calendar()
        else:
            toggle_day_status(date_str, 'error')
            update_calendar()

    toggle_done_btn = tk.Button(status_frame, text="Toggle Done", command=mark_done)
    toggle_done_btn.pack(side=tk.LEFT, padx=5)

    toggle_error_btn = tk.Button(status_frame, text="Toggle Error", command=mark_error)
    toggle_error_btn.pack(side=tk.LEFT, padx=5)

    def draw_calendar(month, year):
        for widget in days_frame.winfo_children():
            widget.destroy()

        month_label.config(text=f"{calendar.month_name[month]} {year}")
        nonlocal cur_month, cur_year
        cur_month, cur_year = month, year

        green_dates = set()

        selected_str = entry.get().strip()
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", selected_str):
            green_dates.add(selected_str)

        if completeAllVar.get():
            try:
                end_m, end_d, end_y = selected_str.split('/')
                end_date = datetime.date(int(end_y), int(end_m), int(end_d))
                expand_db_up_to(end_date)

                conn = sqlite3.connect("days.db")
                c = conn.cursor()
                c.execute("SELECT date FROM days WHERE status != 'done'")
                rows = c.fetchall()
                conn.close()

                for (db_date_str,) in rows:
                    try:
                        mm, dd, yy = db_date_str.split('/')
                        dt = datetime.date(int(yy), int(mm), int(dd))
                        if dt <= end_date:
                            green_dates.add(db_date_str)
                    except:
                        pass
            except:
                pass

        days_of_the_week = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        for col_idx, dow in enumerate(days_of_the_week):
            dow_label = tk.Label(days_frame, text=dow, width=3, font=("Helvetica", 10, "bold"))
            dow_label.grid(row=0, column=col_idx, padx=2, pady=2)

        cal_iter = calendar.Calendar(firstweekday=calendar.SUNDAY)
        for row_idx, week in enumerate(cal_iter.monthdayscalendar(year, month), start=1):
            for col_idx, day in enumerate(week):
                if day == 0:
                    lbl = tk.Label(days_frame, text=" ", width=3)
                    lbl.grid(row=row_idx, column=col_idx)
                else:
                    ds = f"{month:02}/{day:02}/{year}"
                    status = get_day_status(ds)
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
                        entry.delete(0, tk.END)
                        entry.insert(0, ds)
                        update_calendar()

                    btn = tk.Button(days_frame, command=on_day_click, **btn_kwargs)
                    btn.grid(row=row_idx, column=col_idx)

    def update_calendar(event=None):
        text = entry.get().strip()
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
            try:
                new_m, new_d, new_y = text.split('/')
                new_m = int(new_m)
                new_y = int(new_y)
                if 1 <= new_m <= 12:
                    draw_calendar(new_m, new_y)
                else:
                    draw_calendar(cur_month, cur_year)
            except:
                draw_calendar(cur_month, cur_year)
        else:
            draw_calendar(cur_month, cur_year)

    entry.bind("<KeyRelease>", update_calendar)
    draw_calendar(cur_month, cur_year)

    global MAIN_UPDATE_CALENDAR
    MAIN_UPDATE_CALENDAR = update_calendar

    buttons_frame = tk.Frame(window)
    buttons_frame.pack(side=tk.BOTTOM, pady=10)

    def start_button_action():
        if completeAllVar.get():
            threading.Thread(target=complete_range).start()
        else:
            begin(entry.get())

    start_button = tk.Button(buttons_frame, text="Start", width=10, height=2,
                             command=start_button_action)
    start_button.pack(side=tk.LEFT, padx=5)

    pause_button = tk.Button(buttons_frame, text="Pause/Resume", width=15, height=2,
                             command=pause)
    pause_button.pack(side=tk.LEFT, padx=5)

    def on_exit():
        stopNow.set()
        keyboard.unhook_all_hotkeys()
        window.quit()

    exit_button = tk.Button(buttons_frame, text="Exit", command=on_exit,
                            width=10, height=2)
    exit_button.pack(side=tk.LEFT, padx=5)

    window.bind_all("<Tab>", lambda e: window.tk_focusNext().focus())
    window.bind_all("<Return>", lambda e: window.focus_get().invoke())

    window.mainloop()

def safe_update_calendar():
    if root and root.winfo_exists() and MAIN_UPDATE_CALENDAR is not None:
        root.after(1, MAIN_UPDATE_CALENDAR)

def complete_range():
    """
    Find all dates not marked 'done' up to (and including) the date in the Entry.
    For each undone date, run normal logic and mark it done if successful.
    """
    end_str = root.nametowidget(".!toplevel.!frame.!entry").get().strip()
    try:
        end_m, end_d, end_y = end_str.split('/')
        end_date = datetime.date(int(end_y), int(end_m), int(end_d))
    except ValueError:
        print("Invalid date in Entry, cannot complete range.")
        return

    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    c.execute("SELECT date FROM days WHERE status != 'done'")
    rows = c.fetchall()
    conn.close()

    undone_dates = []
    for (date_str,) in rows:
        try:
            m, d, y = date_str.split('/')
            dt = datetime.date(int(y), int(m), int(d))
            if dt <= end_date:
                undone_dates.append((dt, date_str))
        except:
            pass

    undone_dates.sort(key=lambda x: x[0])
    print(f"Completing {len(undone_dates)} dates up to {end_str}...")

    for (dt_obj, d_str) in undone_dates:
        if stopNow.is_set():
            toggle_day_status(d_str, 'error')
            print(f"Stop pressed before processing {d_str}, marking error.")
            continue

        print(f"Processing date {d_str} ...")
        try:
            success = loginToSite(d_str)
            if not success or stopNow.is_set():
                toggle_day_status(d_str, 'error')
                print(f"Marked {d_str} as 'error' (canceled).")
                break
            else:
                toggle_day_status(d_str, 'done')
                print(f"Marked {d_str} as 'done'.")
        except Exception as e:
            toggle_day_status(d_str, 'error')
            print(f"Error on {d_str}: {str(e)}")
    stopNow.clear()
    print("Complete range finished.")

def init_db():
    conn = sqlite3.connect("days.db")
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
    one_day = datetime.timedelta(days=1)
    cur_date = start_date
    while cur_date <= end_date:
        date_str = cur_date.strftime("%m/%d/%Y")
        c.execute("""
            INSERT OR IGNORE INTO days(date, status, error_msg)
            VALUES (?, 'done', '')
        """, (date_str,))
        cur_date += one_day

    conn.commit()
    conn.close()

def get_day_status(date_str):
    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    c.execute("SELECT status FROM days WHERE date=?", (date_str,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else ''

def toggle_day_status(date_str, new_status):
    current = get_day_status(date_str)
    if current == new_status:
        return
    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    c.execute("INSERT OR IGNORE INTO days(date) VALUES(?)", (date_str,))
    c.execute("UPDATE days SET status=? WHERE date=?", (new_status, date_str))
    conn.commit()
    conn.close()

    safe_update_calendar()

def begin(date):
    stopNow.clear()
    threading.Thread(target=loginToSite, args=(date,)).start()

def pause():
    if pauseEvent.is_set():
        pauseEvent.clear()
        print("Paused")
    else:
        pauseEvent.set()
        print("Resumed")

def focus(app):
    app.set_focus()

#######################################################
#               NEW REQUESTS-BASED LOGIC              #
#######################################################

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

    rows = table.find_all("tr", recursive=False)  # top-level <tr> children first
    if not rows:
        # fallback if <tbody> is used
        rows = table.find_all("tr")

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
                    patient_name  = data_tds[0].get_text().rstrip()  # remove trailing space
                    appointment   = data_tds[1].get_text(strip=True)
                    confirmations = data_tds[2].get_text(strip=True)
                    row_status    = data_tds[3].get_text(strip=True)

                    # remove arrow, skip "Family"
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

def loginToSite(date):
    """
    Replaces the old Selenium-based logic with direct requests + AHK typing.
    1) Log in to PracticeMojo
    2) For each cdi/cdn, fetch + parse
    3) Type into SoftDent via AHK
    """
    with open('config.json') as cfg:
        config = json.load(cfg)
    username = config.get('USERNAME')
    password = config.get('PASSWORD')

    try:
        session = login_to_practice_mojo(username, password)
    except Exception as e:
        print("PracticeMojo login failed:", e)
        return False

    # Attempt to connect to SoftDent or WordPad
    app = Application()
    soft = Application()
    try:
        soft.connect(title_re=".*CS SoftDent.*- S.*")
        softdent = soft.window(title_re=".*CS SoftDent.*- S.*")
        print("Connected to SoftDent.")
    except Exception:
        print("SoftDent not found. Connecting to WordPad for testing.")
        soft.connect(title_re=".*WordPad.*")
        softdent = soft.window(title_re=".*WordPad.*")

    cdi_list = [1, 8, 13, 21, 22, 23, 30, 33, 35, 36, 130]
    m, d, y = date.split('/')

    for i in cdi_list:
        for o in range(1, 3):
            if stopNow.is_set():
                toggle_day_status(date, 'error')
                return False

            try:
                data_rows = fetch_activity_detail(session, f"{m}/{d}/{y}", i, o)
            except Exception as exc:
                print(f"Error fetching data for cdi={i}, cdn={o}: {exc}")
                continue

            if not data_rows:
                print(f"No data for cdi={i}, cdn={o}.")
            else:
                print(f"Found {len(data_rows)} row(s) for cdi={i}, cdn={o}.")

            # If cdi in {23, 130}, use nameDateTime logic
            if i in [23, 130]:
                merge_and_type_appointments(data_rows, i, o, m, d, y, softdent, date)
            else:
                type_normal(data_rows, i, o, m, d, y, softdent, date)

    return True

def merge_and_type_appointments(data_rows, cdi, cdn, m, d, y, softdent, date):
    """
    Matches pmojo.py's nameDateTime logic:
    Group all rows by patient, combine times, type 0ca + cc + 'Reminder for' + times, etc.
    """
    # Group times by patient
    grouped = {}
    for row in data_rows:
        patient = row['patient_name'].strip()
        appt = row['appointment']
        if appt:
            grouped.setdefault(patient, []).append(appt)

    typeOfCom = determine_type_of_com(cdi, cdn)
    focus(softdent)

    for patient, times_list in grouped.items():
        pauseEvent.wait()
        if stopNow.is_set():
            return
        # Remove duplicates, sort
        times_str = ' & '.join(sorted(set(times_list)))

        # 1) Tab, 'f', patient, Enter
        ahk.key_press("Tab")
        ahk.type('f')
        pauseEvent.wait()
        if stopNow.is_set():
            return
        ahk.type(patient)
        sleep(2)
        ahk.key_press("Enter")
        pauseEvent.wait()
        if stopNow.is_set():
            return

        # 2) '0ca' + 'cc' => '0cacc', Tab, 'Reminder for ', times_str
        ahk.type("0ca")
        ahk.type("cc")
        ahk.key_press("Tab")
        ahk.type("Reminder for")
        pauseEvent.wait()
        if stopNow.is_set():
            return
        ahk.type(" " + times_str)  # space + merged times

        ahk.key_press("Tab")
        ahk.key_press("Tab")
        ahk.type(typeOfCom)
        ahk.key_press("Tab")
        ahk.type("pmojoNFD")
        ahk.key_press("Tab")
        ahk.type(f"{m}/{d}/{y}")
        sleep(3)
        pauseEvent.wait()
        if stopNow.is_set():
            return
        ahk.key_press("Enter")
        pauseEvent.wait()
        if stopNow.is_set():
            return
        ahk.type("l")
        ahk.type("l")
        ahk.type("n")

def type_normal(data_rows, cdi, cdn, m, d, y, softdent, date):
    """
    Matches pmojo.py's name(...) logic for normal campaigns.
    0ca + optional 'r' + com => typeOfCom => pmojoNFD => date => etc.
    """
    # Determine the 'com' string
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

    typeOfCom = determine_type_of_com(cdi, cdn)
    recare_flag = "r" if cdi in [1, 8, 33] else ""

    focus(softdent)
    for row in data_rows:
        pauseEvent.wait()
        if stopNow.is_set():
            toggle_day_status(date, 'error')
            return

        patient = row['patient_name'].strip()

        # 1) Tab, 'f', patient, Enter
        ahk.key_press("Tab")
        ahk.type('f')
        pauseEvent.wait()
        if stopNow.is_set():
            return
        ahk.type(patient)
        sleep(2)
        ahk.key_press("Enter")
        pauseEvent.wait()
        if stopNow.is_set():
            return

        # 2) "0ca" + recare_flag, Tab, com, Tab, Tab, typeOfCom, ...
        ahk.type("0ca")
        ahk.type(recare_flag)
        ahk.key_press("Tab")
        ahk.type(com)
        pauseEvent.wait()
        if stopNow.is_set():
            return

        ahk.key_press("Tab")
        ahk.key_press("Tab")
        ahk.type(typeOfCom)
        ahk.key_press("Tab")
        ahk.type("pmojoNFD")
        ahk.key_press("Tab")
        ahk.type(f"{m}/{d}/{y}")
        sleep(3)
        pauseEvent.wait()
        if stopNow.is_set():
            return

        ahk.key_press("Enter")
        pauseEvent.wait()
        if stopNow.is_set():
            return

        ahk.type("l")
        ahk.type("l")
        ahk.type("n")

def determine_type_of_com(cdi, cdn):
    """
    Recreates old logic:
      if cdi in [23, 130], cdn=1 => 'e', cdn=2 => 't'
      else name approach:
        cdn=1 => cdi in {1,8,21,22,30,33,35} => 'l', cdi=36 => 'e'
        cdn=2 => cdi in {1,8,30,33} => 'e'
        cdn=3 => cdi=1 => 't'
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

#######################################################
#                 HOTKEYS AND MAIN                    #
#######################################################

keyboard.add_hotkey('num lock', stopProgram)
keyboard.add_hotkey('scroll lock', pause)

if __name__ == "__main__":
    init_db()
    createGUI()
