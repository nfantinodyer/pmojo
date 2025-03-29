from http.server import executable
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
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

warnings.simplefilter('ignore', category=UserWarning)  # Ignore 32bit warnings

ahk = AHK()
pauseEvent = threading.Event()
pauseEvent.set()
root = tk.Tk()
root.withdraw()

def stopProgram():
    print("Stop key pressed. Exiting program...")
    root.quit()

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

    # Find the latest date in the database (if any)
    c.execute("SELECT MAX(date) FROM days")
    row = c.fetchone()
    latest_str = row[0] if row and row[0] else None

    if latest_str:
        # Convert the string to a date
        m, d, y = latest_str.split('/')
        latest_date = datetime.date(int(y), int(m), int(d))
    else:
        # If the DB is empty, let's start from 2020-01-01 or something
        latest_date = datetime.date(2020, 1, 1)

    # If end_date is after the current 'latest_date', fill in the missing days
    if end_date > latest_date:
        one_day = datetime.timedelta(days=1)
        cur_date = latest_date + one_day
        while cur_date <= end_date:
            date_str = cur_date.strftime("%m/%d/%Y")
            c.execute(
                "INSERT OR IGNORE INTO days(date, status, error_msg) VALUES (?, '', '')",
                (date_str,)
            )
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

    # ---------- ENTRY WIDGET -----------
    entry = tk.Entry(frame1, width=50)
    entry.pack(side=tk.LEFT, padx=2, pady=5)
    entry.insert(0, calendar.datetime.date.today().strftime("%m/%d/%Y"))
    entry.focus_set()
    # Pressing Enter on the Entry calls begin(...) in a new thread
    entry.bind("<Return>", lambda e: begin(entry.get()))

    # A BooleanVar to track the "Complete All" checkbox
    completeAllVar = tk.BooleanVar(value=False)

    def on_complete_all_toggle():
        # Whenever the user checks/unchecks "Complete All," redraw the calendar
        update_calendar()

    # Checkbox that triggers a redraw each time it’s toggled
    completeAllCheck = tk.Checkbutton(
        window,
        text="Complete All",
        variable=completeAllVar,
        command=on_complete_all_toggle
    )
    completeAllCheck.pack(anchor=tk.W)

    # ---------- MONTH/YEAR NAVIGATION -----------
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
        # Don't touch the Entry; just redraw
        draw_calendar(cur_month, cur_year)

    def next_month():
        nonlocal cur_month, cur_year
        cur_month += 1
        if cur_month > 12:
            cur_month = 1
            cur_year += 1
        # Don't overwrite the Entry; just redraw
        draw_calendar(cur_month, cur_year)

    left_arrow_btn = tk.Button(month_frame, text="←", command=prev_month)
    left_arrow_btn.pack(side=tk.LEFT, padx=5)

    month_label = tk.Label(month_frame, text="", font=("Helvetica", 12, "bold"))
    month_label.pack(side=tk.LEFT)

    right_arrow_btn = tk.Button(month_frame, text="→", command=next_month)
    right_arrow_btn.pack(side=tk.LEFT, padx=5)

    # ---------- CALENDAR GRID -----------
    days_frame = tk.Frame(window)
    days_frame.pack(side=tk.TOP, anchor="center", padx=2, pady=5)

    # ---------- TOGGLE STATUS BUTTONS -----------
    status_frame = tk.Frame(window)
    status_frame.pack(anchor='center', pady=5)

    selected_day_label = tk.Label(status_frame, text="Status: ")
    selected_day_label.pack(side=tk.LEFT, padx=5)

    def mark_done():
        date_str = entry.get().strip()
        if date_str:
            toggle_day_status(date_str, 'done')
            update_calendar()

    def mark_error():
        date_str = entry.get().strip()
        if date_str:
            toggle_day_status(date_str, 'error')
            update_calendar()

    toggle_done_btn = tk.Button(status_frame, text="Toggle Done", command=mark_done)
    toggle_done_btn.pack(side=tk.LEFT, padx=5)

    toggle_error_btn = tk.Button(status_frame, text="Toggle Error", command=mark_error)
    toggle_error_btn.pack(side=tk.LEFT, padx=5)

    # ---------- DRAW CALENDAR FUNCTION -----------
    def draw_calendar(month, year):
        """
        Redraws the day-buttons for the given month/year.
        1) If "Complete All" is checked, find all undone dates <= the Entry date for green.
        2) Also highlight the current date in the Entry as green (unless it's done/error).
        3) 'done' => gray, 'error' => red take precedence over green.
        """
        # Clear existing day-buttons
        for widget in days_frame.winfo_children():
            widget.destroy()

        # Update the month_label
        month_label.config(text=f"{calendar.month_name[month]} {year}")
        nonlocal cur_month, cur_year
        cur_month, cur_year = month, year

        # We'll store which dates should appear green
        green_dates = set()

        # Always highlight the single date typed in the Entry (unless done/error)
        selected_str = entry.get().strip()
        # Validate the typed date
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", selected_str):
            green_dates.add(selected_str)  # the user's single selection

        # If "Complete All" is checked, gather undone up to that same date
        if completeAllVar.get():
            try:
                end_m, end_d, end_y = selected_str.split('/')
                end_date = datetime.date(int(end_y), int(end_m), int(end_d))

                # 1) Expand DB up to end_date
                expand_db_up_to(end_date)

                # 2) Gather undone dates <= end_date
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
                pass  # user typed something invalid in the Entry

        # Now build the clickable grid
        cal_iter = calendar.Calendar(firstweekday=calendar.SUNDAY)
        for row_idx, week in enumerate(cal_iter.monthdayscalendar(year, month)):
            for col_idx, day in enumerate(week):
                if day == 0:
                    lbl = tk.Label(days_frame, text=" ", width=3)
                    lbl.grid(row=row_idx, column=col_idx)
                else:
                    date_str = f"{month:02}/{day:02}/{year}"
                    status = get_day_status(date_str)

                    bg_color = "SystemButtonFace"
                    if status == "done":
                        bg_color = "gray"
                    elif status == "error":
                        bg_color = "red"
                    # If it's not done/error and is in green_dates => lightgreen
                    elif date_str in green_dates:
                        bg_color = "lightgreen"

                    # Clicking a day sets the Entry to that date and re-draws
                    def on_day_click(ds=date_str):
                        entry.delete(0, tk.END)
                        entry.insert(0, ds)
                        update_calendar()  # so the newly selected date also becomes green

                    btn = tk.Button(
                        days_frame,
                        text=str(day),
                        width=3,
                        bg=bg_color,
                        command=on_day_click
                    )
                    btn.grid(row=row_idx, column=col_idx)

    # ---------- UPDATE CALENDAR ON ENTRY CHANGE -----------
    def update_calendar(event=None):
        """
        Called whenever the user types in the Entry or 
        we programmatically change the Entry (like on_day_click).
        """
        import re
        text = entry.get().strip()
        # If user typed a valid date => re-draw that month
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
            try:
                new_m, new_d, new_y = text.split('/')
                new_m = int(new_m)
                new_y = int(new_y)
                if 1 <= new_m <= 12:
                    draw_calendar(new_m, new_y)
                else:
                    # If out of range (e.g. 13), just re-draw the current
                    draw_calendar(cur_month, cur_year)
            except:
                draw_calendar(cur_month, cur_year)
        else:
            # If invalid, re-draw the current month
            draw_calendar(cur_month, cur_year)

    # Initialize
    # So that if user types something, we automatically re-draw
    entry.bind("<KeyRelease>", update_calendar)
    draw_calendar(cur_month, cur_year)

    # ---------- BOTTOM BUTTONS -----------
    buttons_frame = tk.Frame(window)
    buttons_frame.pack(side=tk.BOTTOM, pady=10)

    def start_button_action():
        if completeAllVar.get():
            threading.Thread(target=complete_range).start()
        else:
            begin(entry.get())

    start_button = tk.Button(
        buttons_frame, text="Start", width=10, height=2,
        command=start_button_action
    )
    start_button.pack(side=tk.LEFT, padx=5)

    pause_button = tk.Button(
        buttons_frame, text="Pause/Resume", width=15, height=2,
        command=pause
    )
    pause_button.pack(side=tk.LEFT, padx=5)

    exit_button = tk.Button(
        buttons_frame, text="Exit", command=window.quit,
        width=10, height=2
    )
    exit_button.pack(side=tk.LEFT, padx=5)

    # Keyboard shortcuts
    window.bind_all("<Tab>", lambda e: window.tk_focusNext().focus())
    window.bind_all("<Return>", lambda e: window.focus_get().invoke())

    window.mainloop()

def complete_range():
    """
    Find all dates not marked 'done' up to (and including) the date in the Entry.
    For each undone date, run our normal logic and mark it done if successful.
    """
    # Parse the target (end) date from the Entry
    end_str = root.nametowidget(".!toplevel.!frame.!entry").get().strip()  
    # ^ The simplest approach is to retrieve the same 'entry' text. If you renamed your Entry widget,
    #   or stored it in a variable, you can just do end_str = entry.get()

    try:
        # Convert end_str ("MM/DD/YYYY") -> Python date
        end_m, end_d, end_y = end_str.split('/')
        end_date = datetime.date(int(end_y), int(end_m), int(end_d))
    except ValueError:
        print("Invalid date in Entry, cannot complete range.")
        return

    # Open DB and find all undone dates that are <= end_date
    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    c.execute("""
        SELECT date
        FROM days
        WHERE status != 'done'
    """)
    rows = c.fetchall()
    conn.close()

    # Convert "MM/DD/YYYY" -> date objects, filter to <= end_date
    undone_dates = []
    for (date_str,) in rows:
        try:
            m, d, y = date_str.split('/')
            dt = datetime.date(int(y), int(m), int(d))
            if dt <= end_date:
                undone_dates.append((dt, date_str))
        except:
            pass

    # Sort ascending
    undone_dates.sort(key=lambda x: x[0])

    print(f"Completing {len(undone_dates)} dates up to {end_str}...")

    # For each undone date, run normal logic:
    for (dt_obj, d_str) in undone_dates:
        if not pauseEvent.is_set():
            # If user clicked Pause, we wait
            pauseEvent.wait()

        print(f"Processing date {d_str} ...")
        try:
            # We reuse your 'loginToSite(...)' logic
            # or do `begin(d_str)` but that spawns a new thread each time;
            # better to just call the loginToSite directly here:
            loginToSite(d_str)

            # If successful, mark as 'done'
            toggle_day_status(d_str, 'done')
            print(f"Marked {d_str} as 'done'.")
        except Exception as e:
            # If there's an error, optionally mark as 'error'
            toggle_day_status(d_str, 'error')
            print(f"Error on {d_str}: {str(e)}")

    print("Complete range finished.")

def init_db():
    """Create or open a local days.db, build the 'days' table, and prefill up to 03/27/2025 as 'done'."""
    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS days(
            date TEXT PRIMARY KEY,
            status TEXT DEFAULT '',   -- 'done', 'error', or ''
            error_msg TEXT DEFAULT ''
        )
    """)
    # We'll fill from some earliest date up to 03/27/2025
    # Adjust 'start_date' to whatever earliest date you want to store
    start_date = datetime.date(2020, 1, 1)
    end_date = datetime.date(2025, 3, 27)

    one_day = datetime.timedelta(days=1)
    cur_date = start_date

    while cur_date <= end_date:
        date_str = cur_date.strftime("%m/%d/%Y")
        # Insert only if not existing
        c.execute("""
            INSERT OR IGNORE INTO days(date, status, error_msg)
            VALUES (?, 'done', '')
        """, (date_str,))
        cur_date += one_day

    conn.commit()
    conn.close()

def get_day_status(date_str):
    """Return the status ('done', 'error', or '') for the given date_str from the DB."""
    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    c.execute("SELECT status FROM days WHERE date=?", (date_str,))
    row = c.fetchone()
    conn.close()
    return row[0] if row else ''

def toggle_day_status(date_str, new_status):
    """
    If the current status is already new_status, toggle it back to ''.
    Otherwise, set it to new_status.
    """
    current = get_day_status(date_str)
    if current == new_status:
        new_status = ''  # revert back
    conn = sqlite3.connect("days.db")
    c = conn.cursor()
    # Insert row if not already present
    c.execute("INSERT OR IGNORE INTO days(date) VALUES(?)", (date_str,))
    # Update status
    c.execute("UPDATE days SET status=? WHERE date=?", (new_status, date_str))
    conn.commit()
    conn.close()

def begin(date):
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

def loginToSite(date):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-notifications")
    driver = webdriver.Chrome(options=chrome_options)

    with open('config.json') as config_file:
        config = json.load(config_file)

    username = config.get('USERNAME')
    password = config.get('PASSWORD')

    #logs into practice mojo
    driver.get("https://app.practicemojo.com/Pages/login")
    elem = driver.find_element(By.NAME, "loginId")
    elem.clear()
    elem.send_keys(username)
    elem = driver.find_element(By.NAME, "password")
    elem.clear()
    elem.send_keys(password)
    elem.send_keys(Keys.RETURN)

    #to be able to set focus with chrome.set_focus() and now focus(chrome)
    app = Application()
    app.connect(title_re='.*- Google Chrome')
    chrome = app.window(title_re='.*- Google Chrome')

    #to be able to set focus with softdent.set_focus() and now focus(softdent)
    soft = Application()
    try:
        soft.connect(title_re=".*CS SoftDent.*- S.*")
        softdent = soft.window(title_re=".*CS SoftDent.*- S.*")
        print("Connected to SoftDent.")
    except Exception as e:
        print("SoftDent not found. Connecting to WordPad for testing.")
        soft.connect(title_re=".*WordPad.*")
        softdent = soft.window(title_re=".*WordPad.*")

    #initial focus set
    focus(chrome)

    # Navigate through the site
    processSite(driver, date, chrome, softdent)
    driver.quit()

def processSite(driver, date, chrome, softdent):
    m, d, y = date.split('/')
    #url
    cdi = [1,8,13,21,22,23,30,33,35,36,130]
    for i in cdi:
        #130 and 23 have appt time and date included
        if(i != 23 and i != 130):
            for o in range(1,3):
                driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi="+str(i)+"&cdn="+str(o))
                #sleep(0.5)
                name(i, o, d, m, y, chrome, softdent)
        else:
            for o in range(1,3):
                driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi="+str(i)+"&cdn="+str(o))
                #sleep(0.5)
                nameDateTime(i, o, d, m, y, chrome, softdent)

def name(cdi, cdn, d, m, y, chrome, softdent):
    #determine if letter email or text based on url
    sleep(0.5)
    typeOfCom = ""
    if cdn == 1:
        if cdi in {1, 8, 21, 22, 30, 33, 35}:
            typeOfCom = "l"
        elif cdi == 36:
            typeOfCom = "e"
    elif cdn == 2:
        if cdi in {1, 8, 30, 33}:
            typeOfCom = "e"
    elif cdn == 3 and cdi == 1:
        typeOfCom = "t"

    # determine communication sent based on url
    com = {1: "Recare: Due", 8: "Recare: Really Past Due", 21: "bday card", 22: "bday card",
           30: "Reactivate: 1 year ago", 33: "Recare: Past Due", 35: "Anniversary card", 36: "bday card"}.get(cdi, "")

    # to copy page text
    send_keys("^a^c")
    clip_text = get_clipboard()

    # create/open txt file to hold copied page contents and paste it in
    with open('practicemojo.txt', "w", encoding="utf-8") as file:
        file.write(clip_text)

    # read the copied text
    with open('practicemojo.txt', 'r') as file:
        text = file.readlines()

    # Remove header and footer lines
    text = text[8:-5]
    unique_lines = set()
    alltext = ""

    for line in text:
        if "Family" in line:  # Skip the line containing "Family" as it's not needed
            continue
        line = re.sub(r"[^\x00-\x7F]+", '', line)  # Remove non-ASCII characters
        line = re.sub(r"\s+", ' ', line).strip()  # Normalize spaces and remove unwanted tabs

        if "Address" in line:
            line = line[:line.find("Address")]  # Trim the line at the word "Address"

        line = line.replace("Bounced", "").replace("Opt Out", "")

        if line not in unique_lines:
            unique_lines.add(line)
            alltext += line + "\n"

    # Write the processed lines to the same file, ensuring there's no extra newline at the end
    with open("practicemojo.txt", "w") as file:
        file.write(alltext.strip())
        
    file.close()

    #softdent

    file1 = open("practicemojo.txt", "r")
    lines = file1.readlines()
    file1.close()

    #brings softdent to the front
    focus(softdent)
    #auto hot key types in softdent and program sleeps after every major input for human to double check
    for line in lines:
        if line != "":
            line = line.strip()

            ahk.key_press("Tab")
            ahk.type('f')
            pauseEvent.wait()
            ahk.type(line)
            sleep(2)
            ahk.key_press('Enter')
            pauseEvent.wait()
            ahk.type("0ca")
            #recare is r
            if cdi == 1 or cdi == 8 or cdi == 33:
                ahk.type("r")
            ahk.key_press("Tab")
            ahk.type(com)
            pauseEvent.wait()
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type(typeOfCom)
            ahk.key_press("Tab")
            ahk.type("pmojoNFD")
            ahk.key_press("Tab")
            
            ahk.type(m+"/"+d+"/"+y)
            sleep(3)
            pauseEvent.wait()
            ahk.key_press("Enter")
            pauseEvent.wait()
            ahk.type("l")
            ahk.type("l")
            ahk.type("n")

    #chrome
    focus(chrome)

def nameDateTime(cdi, cdn, d, m, y, chrome, softdent):
    #determine if letter email or text based on url
    sleep(0.5)
    typeOfCom = ""
    if cdn == 1:
        if cdi == 130 or cdi == 23:
            typeOfCom = "e"
    elif cdn == 2:
        if cdi == 130 or cdi == 23:
            typeOfCom = "t"

    #to copy page text
    send_keys("^a^c")

    clip_text = get_clipboard()

    #create/open txt file to hold copied page contents and paste it in
    go = open('practicemojo.txt' ,"w",encoding="utf-8")
    go.write(clip_text)
    go.close()

    with open('practicemojo.txt', 'r', encoding='utf-8') as text_file:
        text = text_file.readlines()

    # Remove header and footer lines if not part of the needed data
    del text[0:8]
    del text[-5:]

    appointments = {}

    for line in text:
        if re.search("Family",line):  # Skip the line containing "Family" as it's not needed
            continue
        # Normalize spaces and remove non-ASCII characters
        line = re.sub(r"\s+", ' ', line.strip())
        line = re.sub(r"[^\x00-\x7F]+", '', line)

        # Match the expected line format with dates and times
        match = re.search(r"([a-zA-Z, ]+)(\d{2}/\d{2} @ \d{2}:\d{2} [APM]+)", line)
        if match:
            name, datetime = match.groups()
            date, time = datetime.split(' @ ')
            # Construct a unique key for each person and date
            key = f"{name.strip()}{date}"
            if key not in appointments:
                appointments[key] = []
            appointments[key].append(time)
        else:
            print(f"No match for line: {line}")  # Debug output for unmatched lines

    # Prepare to write consolidated entries to file
    lines_to_write = []
    for key, times in appointments.items():
        # Combine multiple times into one line
        times_str = ' & '.join(sorted(set(times)))  # Remove duplicates and sort times
        # Use regex to extract the date part
        match = re.search(r"(\d{2}/\d{2})$", key)
        if match:
            date = match.group(1)
            name = key[:match.start()].strip()
            line_to_write = f"{name} {date} @ {times_str}"
            lines_to_write.append(line_to_write)
        else:
            print(f"Unexpected key format: {key}")  # Debug output for unexpected key format

    # Write all lines to file, avoiding an extra newline on the last line
    with open("practicemojo.txt", "w+", encoding='utf-8') as new_file:
        new_file.write("\n".join(lines_to_write))

    new_file.close()

    #softdent

    #to compare if char is num
    num = "[0-9]+"
    NUMBER = re.compile(num)

    file1 = open("practicemojo.txt", "r")
    lines = file1.readlines()
    file1.close()

    focus(softdent)
    for line in lines:
        now = False
        name = ""
        if line != "":
            for word in line:
                if re.fullmatch(NUMBER, word):
                    break
                else:
                    name+=word

            
            line = line.replace("\n","")
            ahk.key_press("Tab")
            ahk.type('f')
            pauseEvent.wait()
            ahk.type(name)
            sleep(2)
            ahk.key_press("Enter")
            pauseEvent.wait()
            ahk.type("0ca")
            #Confirm Appt is cc
            ahk.type("cc")
            ahk.key_press("Tab")
            ahk.type("Reminder for")
            pauseEvent.wait()
            size = 0
            for word in line:
                size+=1
                if now:
                    line = line[size-1:]
                    now=False
                    break
                elif re.fullmatch(NUMBER, word):
                    ahk.key_press("Space")
                    ahk.type(word)
                    now = True
            
            ahk.type(line)

            pauseEvent.wait()
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type(typeOfCom)
            ahk.key_press("Tab")
            ahk.type("pmojoNFD")
            ahk.key_press("Tab")
            ahk.type(m+"/"+d+"/"+y)
            sleep(3)
            pauseEvent.wait()
            ahk.key_press("Enter")
            pauseEvent.wait()
            ahk.type("l")
            ahk.type("l")
            ahk.type("n")

    #chrome
    focus(chrome)

# Register hotkeys
keyboard.add_hotkey('num lock', stopProgram)
keyboard.add_hotkey('scroll lock', pause)

if __name__ == "__main__":
    init_db()
    createGUI()