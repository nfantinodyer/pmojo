"""SQLite database for tracking which dates have been processed."""
import datetime
import sqlite3


class Database:
    def __init__(self, db_path="days.db", gui=None):
        self.db_path = db_path
        self.gui = gui
        self.init_db()

    @staticmethod
    def _ui_str_to_date(ui_str: str) -> datetime.date:
        """Converts 'MM/DD/YYYY' (UI format) -> datetime.date."""
        return datetime.datetime.strptime(ui_str, "%m/%d/%Y").date()

    @staticmethod
    def _date_to_db_str(d: datetime.date) -> str:
        """Converts datetime.date -> 'YYYY-MM-DD' for storing in DB."""
        return d.strftime("%Y-%m-%d")

    @staticmethod
    def _db_str_to_date(db_str: str) -> datetime.date:
        """Converts 'YYYY-MM-DD' (DB format) -> datetime.date."""
        return datetime.datetime.strptime(db_str, "%Y-%m-%d").date()

    def init_db(self):
        """
        Ensures the 'days' table exists and pre-populates from 2020-01-01
        to 2025-03-27 as 'done', storing each date in 'YYYY-MM-DD' format.
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
        end_date = datetime.date(2025, 3, 27)
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
        """Inserts rows in 'days' table up to 'end_date' if not already present."""
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
        """Return the 'status' field for ui_date_str='MM/DD/YYYY', or '' if not found."""
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

    def toggle_day_status(self, ui_date_str: str, new_status: str):
        """Set the status of the given date. If already set to new_status, do nothing."""
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
