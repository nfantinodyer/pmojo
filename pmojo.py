from http.server import executable
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import tkinter as tk
import re
from ahk import AHK
import time
import json
import warnings

#to ignore pywinauto 32bit warning
warnings.simplefilter('ignore', category=UserWarning)

#auto hot key
ahk = AHK()

#to have a gui to type in date
window = tk.Tk()
window.title("Pmojo")
window.columnconfigure(0, weight=1,minsize=250)
window.rowconfigure(0, weight=1, minsize=250)

#starts program
def begin(date):

    #assign month day and year
    m = date[0:2]
    d = date[3:5]
    y = date[6:10]

    #open chrome with notifs disabled and with chrome driver linked
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

    #to be able to set focus with chrome.set_focus()
    app = Application()
    app.connect(title_re='.*- Google Chrome')
    chrome = app.window(title_re='.*- Google Chrome')

    #initial focus set
    chrome.set_focus()

    #url
    cdi = [1,8,13,21,22,23,30,33,35,36,130]
    for i in cdi:
        #130 and 23 have appt time and date included
        if(i != 23 and i != 130):
            for o in range(1,3):
                driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi="+str(i)+"&cdn="+str(o))
                justName(i,o,d,m,y)
        else:
            for o in range(1,3):
                driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi="+str(i)+"&cdn="+str(o))
                full(i,o,d,m,y)
            

def justName(cdi,cdn,d,m,y):
    #to be able to set focus with chrome.set_focus()
    app = Application()
    app.connect(title_re='.*- Google Chrome')
    chrome = app.window(title_re='.*- Google Chrome')

    #to be able to set focus with softdent.set_focus()
    soft = Application()
    soft.connect(title_re='.*CS SoftDent.*- S')
    softdent = soft.window(title_re='.*CS SoftDent.*- S')

    #determine if letter email or text based on url
    typeOfCom = ""
    if cdn == 1:
        if cdi == 1 or cdi == 8 or cdi == 21 or cdi == 22 or cdi == 30 or cdi == 33 or cdi == 35:
            typeOfCom = "l"
        elif cdi == 36:
            typeOfCom = "e"
    elif cdn == 2:
        if cdi == 1 or cdi == 8 or cdi == 30 or cdi == 33:
            typeOfCom = "e"
    elif cdn == 3:
        if cdi == 1:
            typeOfCom = "t"

    #determine commuication sent based on url
    com = ""
    if cdi == 1:
        com = "Recare: Due"
    elif cdi == 8:
        com = "Recare: Really Past Due"
    elif cdi == 21 or cdi == 22 or cdi == 36:
        com = "bday card"
    elif cdi == 30:
        com = "Reactivate: 1 year ago"
    elif cdi == 33:
        com = "Recare: Past Due"
    elif cdi == 35:
        com = "Anniversary card"

    #to copy page text
    #time.sleep(1)
    send_keys("^a^c")

    #get clipboard
    root = tk.Tk()
    root.withdraw()
    clip_text = root.clipboard_get()

    #create/open txt file to hold copied page contents and paste it in
    go = open('practicemojo.txt' ,"w",encoding="utf-8")
    go.write(clip_text)
    go.close()

    with open('practicemojo.txt', 'r') as text_file:
        text = text_file.readlines()

    # Remove header and footer lines
    del text[0:8]
    del text[-5:]

    unique_lines = set()
    alltext = ""

    for line in text:
        # Normalize the line by removing extra spaces and unwanted tabs
        line = re.sub(r"\s+", ' ', line.strip())

        # Remove any trailing 'Address suppressed' or similar unwanted text
        address_index = line.find("Address")
        if address_index != -1:
            line = line[:address_index]  # Trim the line at the word "Address"

        # Remove non-ASCII characters and specific unwanted phrases
        line = re.sub(r"[^\x00-\x7F]+", '', line)  # Remove non-ASCII characters
        line = line.replace("Bounced", "")  # Remove the word "Bounced"
        line = line.replace("Opt Out", "")  # Remove the phrase "Opt Out"
        line = line.replace("\t", "")  # Remove all tabs

        # Check for duplicates
        if line not in unique_lines:
            unique_lines.add(line)
            alltext += line + "\n"

    # Write the processed lines to the same file, ensuring there's no extra newline at the end
    with open("practicemojo.txt", "w") as new_file:
        new_file.write(alltext.strip())
        
    new_file.close()

    #softdent

    file1 = open("practicemojo.txt", "r")
    lines = file1.readlines()
    file1.close()

    #brings softdent to the front
    softdent.set_focus()

    #auto hot key types in softdent and program sleeps after every major input for human to double check
    for line in lines:
        if line != "":
                
            line = line.replace("\n","")
            ahk.key_press("Tab")
            ahk.type('f')
                
            ahk.type(line)
            time.sleep(1)
            ahk.key_press('Enter')
            ahk.type("0ca")
            #recare is r
            if cdi == 1 or cdi == 8 or cdi == 33:
                ahk.type("r")
            ahk.key_press("Tab")
            ahk.type(com)
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type(typeOfCom)
            ahk.key_press("Tab")
            ahk.type("pmojoNFD")
            ahk.key_press("Tab")
            ahk.type(m+"/"+d+"/"+y)
            time.sleep(2)
                
            ahk.key_press("Enter")

            ahk.type("l")
            ahk.type("l")
            ahk.type("n")

    #chrome
    chrome.set_focus()

def full(cdi,cdn,d,m,y):
    #to be able to set focus with chrome.set_focus()
    app = Application()
    app.connect(title_re='.*- Google Chrome')
    chrome = app.window(title_re='.*- Google Chrome')

    #to be able to set focus with softdent.set_focus()
    soft = Application()
    soft.connect(title_re='.*CS SoftDent.*- S')
    softdent = soft.window(title_re='.*CS SoftDent.*- S')

    #determine if letter email or text based on url
    typeOfCom = ""
    if cdn == 1:
        if cdi == 130 or cdi == 23:
            typeOfCom = "e"
    elif cdn == 2:
        if cdi == 130 or cdi == 23:
            typeOfCom = "t"

    #to copy page text
    send_keys("^a^c")

    #get clipboard
    root = tk.Tk()
    root.withdraw()
    clip_text = root.clipboard_get()

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
        name, date = key.rsplit('10', 1)
        line_to_write = f"{name}10{date} @ {times_str}"
        lines_to_write.append(line_to_write)

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

    softdent.set_focus()

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
                
            ahk.type(name)
            time.sleep(1)
            ahk.key_press("Enter")
            ahk.type("0ca")
            #Confirm Appt is cc
            ahk.type("cc")
            ahk.key_press("Tab")
            ahk.type("Reminder for ")
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
                
            for word in line:
                if re.fullmatch("M", word):
                    ahk.type("M")
                    break
                else:
                    ahk.type(word)
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type(typeOfCom)
            ahk.key_press("Tab")
            ahk.type("pmojoNFD")
            ahk.key_press("Tab")
            ahk.type(m+"/"+d+"/"+y)
            time.sleep(2)
                
            ahk.key_press("Enter")

            ahk.type("l")
            ahk.type("l")
            ahk.type("n")

    #chrome
    chrome.set_focus()


#all for gui
entrylabel = tk.Label(text="Enter the date MM/DD/YYYY: ")
entrylabel.pack(side=tk.TOP,anchor=tk.W)

frame1 = tk.Frame(window, width=50, relief=tk.SUNKEN)
frame1.pack(fill=tk.X)

one = tk.Entry(frame1,width=50)
one.pack(side = tk.LEFT,padx=2,pady=5)

button = tk.Button(text="Start", width=10, height=2, command=lambda: begin(one.get()))
button.pack(anchor=tk.W,padx=2, pady=5)

button = tk.Button(text="Exit", command=quit, width=10, height=2)
button.pack(anchor=tk.W,padx=2, pady=5)

window.mainloop()
