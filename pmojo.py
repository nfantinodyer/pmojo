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

    entrylabel = tk.Label(window, text="Enter the date MM/DD/YYYY: ")
    entrylabel.pack(side=tk.TOP, anchor=tk.W)

    frame1 = tk.Frame(window, width=50, relief=tk.SUNKEN)
    frame1.pack(fill=tk.X)

    entry = tk.Entry(frame1, width=50)
    entry.pack(side=tk.LEFT, padx=2, pady=5)

    start_button = tk.Button(window, text="Start", width=10, height=2, command=lambda: begin(entry.get()))
    start_button.pack(anchor=tk.W, padx=2, pady=5)

    pause_button = tk.Button(window, text="Pause/Resume", width=15, height=2, command=pause)
    pause_button.pack(anchor=tk.W, padx=2, pady=5)

    exit_button = tk.Button(window, text="Exit", command=window.quit, width=10, height=2)
    exit_button.pack(anchor=tk.W, padx=2, pady=5)

    # Bind Tab key to focus the next widget
    window.bind_all("<Tab>", lambda e: window.tk_focusNext().focus())
    # Bind Enter key to activate the focused button
    window.bind_all("<Return>", lambda e: window.focus_get().invoke())

    window.mainloop()

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
            ahk.type("Reminder for ")
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
    createGUI()