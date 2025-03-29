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
    # Pressing Enter in this Entry will call begin(...) with the entry text
    entry.bind("<Return>", lambda e: begin(entry.get()))

    # We store the currently displayed month/year in a mutable list or dict
    # so both arrow buttons and update_calendar can refer to them.
    # Let's initialize from the entry's default date:
    cur_date_str = entry.get()
    m, d, y = cur_date_str.split('/')
    cur_month = int(m)
    cur_year = int(y)
    
    # -- Create a frame to hold the left arrow, month_label, and right arrow in a row --
    month_frame = tk.Frame(window)
    month_frame.pack(pady=(5, 0), anchor="center")

    # Left arrow button
    def prev_month():
        nonlocal cur_month, cur_year
        cur_month -= 1
        if cur_month < 1:
            cur_month = 12
            cur_year -= 1
        # For simplicity, let's just set the day to "1"
        new_date_str = f"{cur_month:02}/01/{cur_year}"
        entry.delete(0, tk.END)
        entry.insert(0, new_date_str)
        draw_calendar(cur_month, cur_year)

    left_arrow_btn = tk.Button(
        month_frame,
        text="←",
        command=prev_month
    )
    left_arrow_btn.pack(side=tk.LEFT, padx=5)

    # Label that shows "March 2025", etc. -- center it in month_frame
    month_label = tk.Label(month_frame, text="", font=("Helvetica", 12, "bold"))
    month_label.pack(side=tk.LEFT)

    # Right arrow button
    def next_month():
        nonlocal cur_month, cur_year
        cur_month += 1
        if cur_month > 12:
            cur_month = 1
            cur_year += 1
        # For simplicity, let's set day to "1"
        new_date_str = f"{cur_month:02}/01/{cur_year}"
        entry.delete(0, tk.END)
        entry.insert(0, new_date_str)
        draw_calendar(cur_month, cur_year)

    right_arrow_btn = tk.Button(
        month_frame,
        text="→",
        command=next_month
    )
    right_arrow_btn.pack(side=tk.LEFT, padx=5)

    # Frame that holds the clickable day-buttons
    days_frame = tk.Frame(window)
    days_frame.pack(side=tk.TOP, anchor="center", padx=2, pady=5)

    # Frame for toggling status
    status_frame = tk.Frame(window)
    status_frame.pack(anchor='center', pady=5)

    # A label to show the selected day from the Entry
    selected_day_label = tk.Label(status_frame, text="Selected: ")
    selected_day_label.pack(side=tk.LEFT, padx=5)

    def mark_done():
        """Toggle the selected date to 'done' or back to ''."""
        date_str = entry.get().strip()
        if date_str:
            toggle_day_status(date_str, 'done')
            update_calendar()  # refresh button colors

    def mark_error():
        """Toggle the selected date to 'error' or back to ''."""
        date_str = entry.get().strip()
        if date_str:
            toggle_day_status(date_str, 'error')
            update_calendar()  # refresh button colors

    toggle_done_btn = tk.Button(status_frame, text="Toggle Done", command=mark_done)
    toggle_done_btn.pack(side=tk.LEFT, padx=5)

    toggle_error_btn = tk.Button(status_frame, text="Toggle Error", command=mark_error)
    toggle_error_btn.pack(side=tk.LEFT, padx=5)

    def draw_calendar(month, year):
        """Rebuilds the day-buttons for the given month/year, centered under the label."""
        # Clear out old buttons
        for widget in days_frame.winfo_children():
            widget.destroy()

        # Update the month_label to something like "March 2025"
        month_label.config(text=f"{calendar.month_name[month]} {year}")

        # Also update our stored month/year so arrows remain in sync
        nonlocal cur_month, cur_year
        cur_month, cur_year = month, year

        cal_iter = calendar.Calendar(firstweekday=calendar.SUNDAY)
        # Build the clickable grid of days
        for row_idx, week in enumerate(cal_iter.monthdayscalendar(year, month)):
            for col_idx, day in enumerate(week):
                if day == 0:
                    # 0 means day is from a previous/next month
                    lbl = tk.Label(days_frame, text=" ", width=3)
                    lbl.grid(row=row_idx, column=col_idx)
                else:
                    date_str = f"{month:02}/{day:02}/{year}"
                    # Get the status for this date from the DB
                    status = get_day_status(date_str)

                    # Decide background color
                    bg_color = "SystemButtonFace"  # default
                    if status == "done":
                        bg_color = "gray"
                    elif status == "error":
                        bg_color = "red"

                    btn = tk.Button(
                        days_frame,
                        text=str(day),
                        width=3,
                        bg=bg_color,
                        command=lambda ds=date_str: (
                            entry.delete(0, tk.END),
                            entry.insert(0, ds)
                        )
                    )
                    btn.grid(row=row_idx, column=col_idx)

    def update_calendar(event=None):
        """Try to parse the entry as MM/DD/YYYY, redraw the days if valid."""
        import re
        text = entry.get().strip()
        # Quick check: must match "MM/DD/YYYY"
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
            try:
                new_m, new_d, new_y = text.split('/')
                new_m = int(new_m)
                new_y = int(new_y)
                if 1 <= new_m <= 12:
                    draw_calendar(new_m, new_y)
            except ValueError:
                pass  # if anything fails, ignore

    # Whenever the user types in the Entry, try to update the calendar
    entry.bind("<KeyRelease>", update_calendar)

    # Draw the initial calendar for today's date
    draw_calendar(cur_month, cur_year)

    # ---- Buttons at the bottom in a single row ----
    buttons_frame = tk.Frame(window)
    buttons_frame.pack(side=tk.BOTTOM, pady=10)

    start_button = tk.Button(
        buttons_frame, text="Start", width=10, height=2,
        command=lambda: begin(entry.get())
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

    # Bind Tab key to focus the next widget
    window.bind_all("<Tab>", lambda e: window.tk_focusNext().focus())
    # Bind Enter key to invoke whichever widget is focused
    window.bind_all("<Return>", lambda e: window.focus_get().invoke())

    window.mainloop()

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
    soft.connect(title_re='.*CS SoftDent.*- S.*')
    softdent = soft.window(title_re='.*CS SoftDent.*- S.*')
    
    #for testing
    #soft.connect(title_re='.*WordPad.*')
    #softdent = soft.window(title_re='.*WordPad.*')

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