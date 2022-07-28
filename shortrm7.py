from binhex import LINELEN
from http.server import executable
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
import os
import os.path
import tkinter as tk
import re
from ahk import AHK
import time
from selenium.webdriver.chrome.service import Service

#auto hot key
ahk = AHK()

#to open a chrome instance
service = Service(executable_path="C:\\chromedriver.exe")

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
    driver = webdriver.Chrome(options=chrome_options, service=service)

    #logs into practice mojo
    driver.get("http://www.practicemojo.com/login")
    elem = driver.find_element(By.NAME, "loginId")
    elem.clear()
    elem.send_keys("oakmeadowdental")
    elem = driver.find_element(By.NAME, "password")
    elem.clear()
    elem.send_keys("OMDC20162017!")
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
    soft.connect(title_re='.*- S')
    softdent = soft.window(title_re='.*- S')

    #determine if letter email or text based on url
    typeOfCom = ""
    if cdn == 1:
        if cdi == 1 or cdi == 8 or cdi == 21 or cdi == 22 or cdi == 30 or cdi == 33 or cdi == 35 or cdi == 36:
            typeOfCom = "l"
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
    time.sleep(1)
    send_keys("^a^c")

    #get clipboard
    root = tk.Tk()
    root.withdraw()
    clip_text = root.clipboard_get()

    #create/open txt file to hold copied page contents and paste it in
    go = open('practicemojo.txt' ,"w",encoding="utf-8")
    go.write(clip_text)
    go.close()

    #read lines of the file
    text_file = open('practicemojo.txt', 'r')
    text = text_file.readlines()
    text_file.close()

    #remove extra parts on the page
    del text[0:7+1]
    del text[-5:]

    #logic for getting pt names out of the file
    new_file = open("practicemojo.txt", "w+")
    temp = text
    skip5=0
    alltext=""
    single=""
    lastlength = 1
    totallines = 0

    for line in temp:
        totallines += 1

        #remove address supressed at the end of the line as well as making sure that it doesnt repeat pts as pmojo sometimes prints families multiple times.
        if line[0:lastlength+3] not in single:
            single += line
            if re.search("Family", line):
                continue
            if re.search("Address",line):
                lengthRemove = 0
                for i in range(5, len(line)):
                    if line[i] == "A" and line[i+1] == "d" and line[i+2] == "d" and line[i+3] == "r" and line[i+4] == "e" and line[i+5] == "s":
                        lengthRemove = i
                        break
                line = line[0:lengthRemove]
                line += "\n"

            count=0
            line = line.replace("\t","")

            #removes special chars after eat line of "Family" and adds line to alltext
            for letter in line:
                count=count+1
                if skip5>0:
                    skip5 -= 1
                elif letter == " " and count==1:
                    skip5 = 5
                elif letter == ",":
                    lastlength = count
                    alltext += letter
                else:
                    alltext += letter
       
        alltext = alltext.rstrip()
        alltext+="\n"
    
    #removes end of line text
    alltext = alltext[:-1]
    alltext = alltext.replace("\t","")
    alltext = alltext.replace("Bounced","")
    alltext = alltext.replace("Opt Out","")
    alltext = alltext.strip()

    #write to new_file
    for line in alltext:
        new_file.write(line)
        
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
            time.sleep(3)
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
    soft.connect(title_re='.*- S')
    softdent = soft.window(title_re='.*- S')

    #determine if letter email or text based on url
    typeOfCom = ""
    if cdn == 1:
        if cdi == 130 or cdi == 23:
            typeOfCom = "e"
    elif cdn == 2:
        if cdi == 130 or cdi == 23:
            typeOfCom = "t"

    #to copy page text
    time.sleep(1)
    send_keys("^a^c")

    #get clipboard
    root = tk.Tk()
    root.withdraw()
    clip_text = root.clipboard_get()

    #create/open txt file to hold copied page contents and paste it in
    go = open('practicemojo.txt' ,"w",encoding="utf-8")
    go.write(clip_text)
    go.close()

    #read lines of the file
    text_file = open('practicemojo.txt', 'r')
    text = text_file.readlines()
    text_file.close()

    #remove extra parts on the page
    del text[0:7+1]
    del text[-5:]

    #logic to get names dates and times from file
    new_file = open("practicemojo.txt", "w+")
    temp = text
    skip5=0
    next6=0
    alltext=""
    single=""
    liststring=[]
    lastlength = 1
    totallines = 0

    for line in temp:
        totallines += 1
        if line[0:lastlength+3] in single:
            if re.search("@", line):
                #puts first time first no matter the order in practice mojo. Removes second occurance.
                for li in liststring:
                    if li[0:lastlength+3]==line[0:lastlength+3]:
                        #since 12pm is earlier than 3pm but 12>3 so this corrects that and fixes 7am time just in case
                        if line[next6+1:next6+3]<="12" and line[next6+1:next6+3]>="7":
                            if line[next6+1:next6+3] > li[next6+1:next6+3]:
                                single = single[0:len(single)-len(li)]
                                alltext = alltext[0:len(alltext)-len(li)]
                        #since 3pm is before 4pm and 3<4 
                        elif line[next6+1:next6+3] < li[next6+1:next6+3]:
                            single = single[0:len(single)-len(li)]
                            alltext = alltext[0:len(alltext)-len(li)]    

        #skips family and gets time and date after the @ sign.
        if line[0:lastlength+3] not in single:
            single+=line
            liststring.append(line)
            if re.search("Family", line):
                continue
            if re.search("@", line):
                count=0
                for letter in line:
                    count=count+1
                    if skip5>0:
                        skip5 -= 1
                    elif letter == " " and count==1:
                        skip5 = 5
                    elif letter == ",":
                        lastlength = count
                        alltext += letter
                    elif letter == "@":
                        alltext += "@"
                        next6=count
                    else:
                        alltext += letter


        alltext = alltext.rstrip()
        alltext+="\n"
        
    #dont need to replace opt out and address supressed due to the next6 being the next 6 chars 00:00AM
    alltext = alltext[:-1]
    alltext = alltext.replace("\t","")
    alltext = alltext.strip()

    for line in alltext:
        new_file.write(line)

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
            time.sleep(3)
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