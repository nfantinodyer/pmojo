import threading
import tkinter as tk
import time
import os

# Event for pausing/unpausing the thread
pause_event = threading.Event()
pause_event.set()  # Start in un-paused state

# Define your functions
def start():
    thread.start()

def pause():
    pause_event.clear()  # Pause the thread

def resume():
    pause_event.set()  # Resume the thread

def exit():
    os._exit(0)

def begin():
    # Use self.pause_event.wait() instead of time.sleep()
    while True:
        print("Working...")
        while not pause_event.is_set():
            time.sleep(0.1)

# Create your thread
thread = threading.Thread(target=begin)

root = tk.Tk()

# Create your buttons
start_button = tk.Button(text="Start", command=start)
start_button.pack()

pause_button = tk.Button(text="Pause", command=pause)
pause_button.pack()

resume_button = tk.Button(text="Resume", command=resume)
resume_button.pack()

exit_button = tk.Button(text="Exit", command=exit)
exit_button.pack()

root.mainloop()
