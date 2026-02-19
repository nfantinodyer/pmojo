"""Tkinter GUI for pmojo with calendar, date selection, and status management."""
import re
import json
import threading
import calendar
import datetime
import sqlite3
import platform
import tkinter as tk
from tkinter import messagebox

if platform.system() == "Windows":
    import keyboard

from pmojo_lib.events import STOP_EVENT, PAUSE_EVENT
from pmojo_lib.database import Database
from pmojo_lib.practicemojo_api import PracticeMojoAPI
from pmojo_lib.softdent import SoftDentConnection
from pmojo_lib.automation import PmojoAutomation, ParseError


class PmojoGUI:
    def __init__(self, db: Database):
        self.db = db
        self.pm_api = None
        self.root = tk.Tk()
        self.root.withdraw()
        self.window = tk.Toplevel(self.root)
        self.window.title("Pmojo")
        self.completeAllVar = tk.BooleanVar(value=False)
        self.entry = None
        self.cur_month = 0
        self.cur_year = 0
        self.create_gui()
        self._login_at_startup()

    def _login_at_startup(self):
        """Login to PracticeMojo in background at app launch so it's ready when user hits Start."""
        def do_login():
            try:
                with open("config.json") as cfg:
                    conf = json.load(cfg)
                pm_api = PracticeMojoAPI(conf["USERNAME"], conf["PASSWORD"])
                pm_api.login()
                self.pm_api = pm_api
                print("Logged into PracticeMojo (startup).")
            except Exception as e:
                print(f"Startup login failed (will retry when Start is pressed): {e}")
        threading.Thread(target=do_login, daemon=True).start()

    def run(self):
        self.window.mainloop()

    def create_gui(self):
        self.is_running = False

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

        default_date = datetime.date.today().strftime("%m/%d/%Y")
        self.entry = tk.Entry(frame1, width=50)
        self.entry.pack(side=tk.LEFT, padx=2, pady=5)
        self.entry.insert(0, default_date)
        self.entry.focus_set()
        self.entry.bind("<Return>", lambda e: self.begin(self.entry.get()))

        self.chk_complete_all = tk.Checkbutton(
            self.window,
            text="Complete All",
            variable=self.completeAllVar,
            command=self.update_calendar
        )
        self.chk_complete_all.pack(anchor=tk.W)

        m, d, y = default_date.split('/')
        self.cur_month = int(m)
        self.cur_year = int(y)

        month_frame = tk.Frame(self.window)
        month_frame.pack(pady=(5, 0), anchor="center")

        self.btn_left = tk.Button(month_frame, text="\u2190", command=self.prev_month)
        self.btn_left.pack(side=tk.LEFT, padx=5)

        self.month_label = tk.Label(month_frame, text="", font=("Helvetica", 12, "bold"))
        self.month_label.pack(side=tk.LEFT)

        self.btn_right = tk.Button(month_frame, text="\u2192", command=self.next_month)
        self.btn_right.pack(side=tk.LEFT, padx=5)

        self.days_frame = tk.Frame(self.window)
        self.days_frame.pack(side=tk.TOP, anchor="center", padx=2, pady=5)

        status_frame = tk.Frame(self.window)
        status_frame.pack(anchor='center', pady=5)

        lbl_status = tk.Label(status_frame, text="Status: ")
        lbl_status.pack(side=tk.LEFT, padx=5)

        self.btn_done = tk.Button(status_frame, text="Toggle Done", command=self.mark_done)
        self.btn_done.pack(side=tk.LEFT, padx=5)

        self.btn_error = tk.Button(status_frame, text="Toggle Error", command=self.mark_error)
        self.btn_error.pack(side=tk.LEFT, padx=5)

        self.entry.bind("<KeyRelease>", self.update_calendar)
        self.draw_calendar(self.cur_month, self.cur_year)

        btns_frame = tk.Frame(self.window)
        btns_frame.pack(side=tk.BOTTOM, pady=10)

        tips_frame = tk.LabelFrame(btns_frame, text="Keybinds", bd=2, relief=tk.GROOVE, font=("Helvetica", 9, "bold"))
        tips_frame.pack(side=tk.LEFT, padx=5, pady=5)
        tips_label = tk.Label(
            tips_frame,
            text="Scroll Lock = Pause/Resume\nNum Lock = Stop Program",
            anchor="w",
            justify="left",
        )
        tips_label.pack(side=tk.LEFT, padx=5)

        self.btn_start = tk.Button(btns_frame, text="Start", width=10, height=2, command=self.start_button_action)
        self.btn_start.pack(side=tk.LEFT, padx=5)

        self.btn_exit = tk.Button(btns_frame, text="Exit", width=10, height=2, command=self.on_exit)
        self.btn_exit.pack(side=tk.LEFT, padx=5)

        self.window.bind_all("<Tab>", lambda e: self.window.tk_focusNext().focus())
        self.window.bind_all("<Return>", lambda e: self.window.focus_get().invoke())

        self.update_calendar()

    def set_running(self, running: bool):
        """Disable/enable UI controls during processing."""
        self.is_running = running
        new_state = "disabled" if running else "normal"

        if self.btn_start:
            self.btn_start.config(state=new_state)
        if self.btn_done:
            self.btn_done.config(state=new_state)
        if self.btn_error:
            self.btn_error.config(state=new_state)
        if self.btn_left:
            self.btn_left.config(state=new_state)
        if self.btn_right:
            self.btn_right.config(state=new_state)
        if self.chk_complete_all:
            self.chk_complete_all.config(state="disabled" if running else "normal")
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

    def _day_color(self, status, is_green):
        """Return the color for a calendar day button.

        On macOS, tk.Button ignores 'bg' — only 'highlightbackground' works,
        so we use that on Mac and 'bg' on Windows.
        """
        if status == "done":
            return "gray"
        elif status == "error" and is_green:
            return "green"
        elif status == "error":
            return "red"
        elif is_green:
            return "lightgreen"
        else:
            return "SystemButtonFace"

    def draw_calendar(self, month, year):
        for widget in self.days_frame.winfo_children():
            widget.destroy()

        self.month_label.config(text=f"{calendar.month_name[month]} {year}")

        text = self.entry.get().strip()
        green_dates = set()

        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", text):
            green_dates.add(text)

        if self.completeAllVar.get():
            try:
                mm, dd, yy = text.split('/')
                end_date = datetime.date(int(yy), int(mm), int(dd))
                self.db.expand_db_up_to(end_date)

                conn = sqlite3.connect(self.db.db_path)
                c = conn.cursor()
                c.execute("SELECT date FROM days WHERE status != 'done'")
                rows = c.fetchall()
                conn.close()

                for (db_date_str,) in rows:
                    try:
                        dt = self.db._db_str_to_date(db_date_str)
                        if dt <= end_date:
                            ui_str = dt.strftime("%m/%d/%Y")
                            green_dates.add(ui_str)
                    except Exception:
                        pass
            except Exception:
                pass

        is_mac = platform.system() == "Darwin"
        # On macOS, tk.Button ignores 'bg'; use 'highlightbackground' instead
        color_key = "highlightbackground" if is_mac else "bg"

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
                    tk.Label(self.days_frame, text=" ", width=3).grid(row=row_idx, column=col_idx)
                else:
                    ds = f"{month:02}/{day:02}/{year}"
                    status = self.db.get_day_status(ds)
                    color = self._day_color(status, ds in green_dates)
                    btn_kwargs = {"text": str(day), "width": 3, color_key: color}

                    # On Windows, also set highlightbackground for error+green combo
                    if not is_mac and status == "error" and ds in green_dates:
                        btn_kwargs["highlightbackground"] = "red"

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
            except Exception:
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
        if platform.system() == "Windows":
            keyboard.unhook_all_hotkeys()
        self.window.quit()

    def _get_pm_api(self):
        """Return cached PM API session, or login fresh if needed."""
        if self.pm_api and self.pm_api.session:
            return self.pm_api
        print("Logging into PracticeMojo...")
        with open("config.json") as cfg:
            conf = json.load(cfg)
        pm_api = PracticeMojoAPI(conf["USERNAME"], conf["PASSWORD"])
        pm_api.login()
        self.pm_api = pm_api
        print("Logged into PracticeMojo.")
        return pm_api

    def begin(self, date_str: str):
        STOP_EVENT.clear()

        def background():
            try:
                try:
                    pm_api = self._get_pm_api()
                except Exception as e:
                    print(f"PracticeMojo login failed: {e}")
                    self.db.toggle_day_status(date_str, "error")
                    return
                automator = SoftDentConnection()
                pm = PmojoAutomation(self.db, pm_api, automator)
                pm.process_date(date_str)
            except ParseError as e:
                STOP_EVENT.set()
                self.db.toggle_day_status(date_str, "error")
                print(f"PARSE ERROR on {date_str}: {e}")
                self.window.after(0, lambda: messagebox.showerror(
                    "Parse Error - Stopped",
                    f"Automation stopped due to a parse error on {date_str}:\n\n{e}\n\n"
                    "No data was typed for this row. Please check PracticeMojo's HTML format."
                ))
            except Exception as e:
                self.db.toggle_day_status(date_str, "error")
                print(f"Unexpected error on {date_str}: {e}")
            finally:
                self.window.after(0, lambda: self.set_running(False))

        threading.Thread(target=background).start()

    def complete_range(self):
        end_str = self.entry.get().strip()
        try:
            mm, dd, yy = end_str.split('/')
            end_date = datetime.date(int(yy), int(mm), int(dd))
        except Exception:
            print("Invalid date in Entry, cannot complete range.")
            return

        conn = sqlite3.connect(self.db.db_path)
        c = conn.cursor()
        c.execute("SELECT date FROM days WHERE status != 'done'")
        rows = c.fetchall()
        conn.close()

        undone = []
        for (db_date_str,) in rows:
            try:
                dt = self.db._db_str_to_date(db_date_str)
                if dt <= end_date:
                    ds_ui = dt.strftime("%m/%d/%Y")
                    undone.append((dt, ds_ui))
            except Exception:
                pass

        undone.sort(key=lambda x: x[0])
        print(f"Completing {len(undone)} dates up to {end_str}...")

        def background():
            try:
                STOP_EVENT.clear()
                try:
                    pm_api = self._get_pm_api()
                except Exception as e:
                    print(f"PracticeMojo login failed: {e}")
                    return
                automator = SoftDentConnection()
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
                    except ParseError as e:
                        STOP_EVENT.set()
                        self.db.toggle_day_status(ds_ui, "error")
                        print(f"PARSE ERROR on {ds_ui}: {e}")
                        self.window.after(0, lambda: messagebox.showerror(
                            "Parse Error - Stopped",
                            f"Automation stopped due to a parse error on {ds_ui}:\n\n{e}\n\n"
                            "No data was typed for this row. Please check PracticeMojo's HTML format."
                        ))
                        return
                    except Exception as ex:
                        self.db.toggle_day_status(ds_ui, "error")
                        print(f"Error on {ds_ui}: {ex}")
                STOP_EVENT.clear()
                print("Complete range finished.")
            finally:
                self.window.after(0, lambda: self.set_running(False))

        threading.Thread(target=background).start()
