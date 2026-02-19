"""
pmojo - Automates logging PracticeMojo campaign data into SoftDent.

This is the entry point. All logic lives in pmojo_lib/.
"""
import platform
import warnings

warnings.simplefilter("ignore", category=UserWarning)

from pmojo_lib.events import STOP_EVENT, PAUSE_EVENT
from pmojo_lib.database import Database
from pmojo_lib.gui import PmojoGUI


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


if platform.system() == "Windows":
    import keyboard
    keyboard.add_hotkey("num lock", stop_program)
    keyboard.add_hotkey("scroll lock", pause_program)

if __name__ == "__main__":
    db = Database("days.db")
    gui = PmojoGUI(db)
    db.gui = gui
    gui.run()
