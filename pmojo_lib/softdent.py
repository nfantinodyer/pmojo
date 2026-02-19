"""SoftDent window automation via pywinctl + AHK (Windows) or AppleScript/pyautogui (Mac)."""
import platform
import re
import subprocess


class _MacTyper:
    """AHK-compatible typing interface using pyautogui on Mac."""

    KEY_MAP = {
        "Tab": "tab",
        "Enter": "enter",
        "Escape": "escape",
        "Space": "space",
    }

    def __init__(self):
        import pyautogui
        self._pyautogui = pyautogui
        self._pyautogui.PAUSE = 0.02

    def type(self, text: str):
        self._pyautogui.typewrite(text, interval=0.02)

    def key_press(self, key: str):
        mapped = self.KEY_MAP.get(key, key.lower())
        self._pyautogui.press(mapped)


class SoftDentConnection:
    """Connect to SoftDent (or fallback test app) and handle typing."""

    def __init__(self):
        self.win = None
        self._mac_app = None  # Store Mac app name for AppleScript activation
        if platform.system() == "Windows":
            from ahk import AHK
            self.ahk = AHK()
        else:
            self.ahk = _MacTyper()

    def connect(self):
        """Connect to SoftDent, else fallback to WordPad/TextEdit."""
        if platform.system() == "Darwin":
            self._connect_mac()
        else:
            self._connect_windows()

    def _connect_mac(self):
        """Use AppleScript to find TextEdit (no accessibility permissions needed)."""
        try:
            result = subprocess.run(
                ["osascript", "-e",
                 'tell application "System Events" to get name of every process whose visible is true'],
                capture_output=True, text=True, timeout=5
            )
            visible_apps = [a.strip() for a in result.stdout.strip().split(",")]

            if "TextEdit" in visible_apps:
                self._mac_app = "TextEdit"
                print("SoftDent not found. Connected to TextEdit for testing.")
            else:
                print("TextEdit not found. Open TextEdit to test.")
        except Exception as e:
            print(f"Could not find windows on Mac: {e}")

    def _connect_windows(self):
        """Use pywinctl to find SoftDent or WordPad on Windows."""
        import pywinctl

        pattern = re.compile(r".*CS SoftDent.*- S.*")
        all_windows = pywinctl.getAllWindows()

        for w in all_windows:
            if pattern.match(w.title):
                self.win = w
                print("Connected to SoftDent.")
                return

        for w in all_windows:
            if "WordPad" in w.title:
                self.win = w
                print("SoftDent not found. Connected to WordPad for testing.")
                return

        print("Neither SoftDent nor WordPad found.")

    def focus(self):
        if platform.system() == "Darwin" and self._mac_app:
            subprocess.run(
                ["osascript", "-e",
                 f'tell application "{self._mac_app}" to activate'],
                capture_output=True, timeout=5
            )
        elif self.win is not None:
            self.win.activate()

    def type_text(self, text: str):
        if self.ahk:
            self.ahk.type(text)

    def press_tab(self, times=1):
        if self.ahk:
            for _ in range(times):
                self.ahk.key_press("Tab")
