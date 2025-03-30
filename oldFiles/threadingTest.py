import threading
import tkinter as tk
import time
import os

class Application:
    def __init__(self, master):
        self.master = master
        self.pause_event = threading.Event()
        self.pause_event.set()  # Start in un-paused state
        self.thread = threading.Thread(target=self.begin)
        
        self.entry = tk.Entry(master)
        self.entry.pack()
        
        self.start_button = tk.Button(master, text="Start", command=self.start)
        self.start_button.pack()
        
        self.pause_button = tk.Button(master, text="Pause", command=self.pause)
        self.pause_button.pack()
        
        self.resume_button = tk.Button(master, text="Resume", command=self.resume)
        self.resume_button.pack()
        
        self.exit_button = tk.Button(master, text="Exit", command=self.exit)
        self.exit_button.pack()

    def start(self):
        date = self.entry.get()
        self.thread.start()

    def pause(self):
        self.pause_event.clear()  # Pause the thread

    def resume(self):
        self.pause_event.set()  # Resume the thread

    def exit(self):
        os._exit(0)

    def begin(self):
        # Use self.pause_event.wait() instead of time.sleep()
        while True:
            print("Working...")
            while not self.pause_event.is_set():
                time.sleep(0.1)

root = tk.Tk()
app = Application(root)
root.mainloop()