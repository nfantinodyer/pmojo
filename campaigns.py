from binhex import LINELEN
from http.server import executable
from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from datetime import date
from pywinauto.application import Application
import pywinauto.mouse as mouse
import pywinauto.keyboard as keyboard
from pywinauto.keyboard import send_keys
import os
import os.path
import shutil
import tkinter as tk
import re
from ahk import AHK
import time

window = tk.Tk()
window.title("Pmojo")
window.columnconfigure(0, weight=1,minsize=250)
window.rowconfigure(0, weight=1, minsize=250)

def begin(date):

    m = date[0:2]
    d = date[3:5]
    y = date[6:10]

    #ahk
    ahk = AHK(executable_path="C:\\Users\\T1.OMDC\\AppData\\Local\\Programs\\Python\\Python310\\Scripts\\AutoHotKey.exe")

    #chrome
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-notifications")
    driver = webdriver.Chrome(options=chrome_options)
    #driver = webdriver.Chrome() 
    driver.get("http://www.practicemojo.com/login")
    elem = driver.find_element(By.NAME, "loginId")
    elem.clear()
    elem.send_keys("oakmeadowdental")
    elem = driver.find_element(By.NAME, "password")
    elem.clear()
    elem.send_keys("OMDC20162017!")
    elem.send_keys(Keys.RETURN)

    app = Application()
    app.connect(title_re='.*- Google Chrome')
    chrome = app.window(title_re='.*- Google Chrome')
    #chrome.set_focus()

    soft = Application()
    soft.connect(title_re='.*- S')
    softdent = soft.window(title_re='.*- S')
    #softdent.set_focus()

    #chrome
    chrome.set_focus()

    #email
    driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=8&cdn=1")#debretire email
    time.sleep(1)
    send_keys("^a^c")


    root = tk.Tk()
    root.withdraw()
    clip_text = root.clipboard_get()

    go = open('debretire.txt' ,"w",encoding="utf-8")
    go.write(clip_text)
    go.close()


    text_file = open('debretire.txt', 'r')
    text = text_file.readlines()
    text_file.close()

    del text[0:7+1]
    del text[-5:]

    new_file = open("debretire.txt", "w+")

    temp = text
    skip5=0

    alltext=""
    single=""
    lastlength = 1

    totallines = 0
    for line in temp:
        totallines += 1
        numcomma = 0
        commaindex = 0
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
            for letter in line:
                count=count+1
                if skip5>0:
                    skip5 -= 1
                elif letter == " " and count==1:
                    skip5 = 5
                elif letter == ",":
                    numcomma+=1
                    if numcomma==1:
                        commaindex=len(alltext)
                    lastlength = count
                    alltext += letter
                else:
                    alltext += letter
        #if second comma (Phd or DDS or etc removes first comma)
        if numcomma>1:
            tempt=""
            noti = commaindex
            for i in range(len(alltext)):
                if i != noti:
                    tempt = tempt + alltext[i]
            alltext = tempt
        alltext = alltext.rstrip()
        alltext+="\n"
        
    alltext = alltext[:-1]
    alltext = alltext.replace("\t","")
    alltext = alltext.replace("Bounced","")
    alltext = alltext.replace("Opt Out","")
    alltext = alltext.strip()

    for line in alltext:
        new_file.write(line)
        

    new_file.close()

    #softdent
    #app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

    num = "[0-9]+"

    NUMBER = re.compile(num)

    file1 = open("recarereallyL.txt", "r")
    lines = file1.readlines()
    file1.close()

    softdent.set_focus()


    for line in lines:
        if line != "":
                
            line = line.replace("\n","")
            ahk.key_press("Tab")
            ahk.type('f')
                
            ahk.type(line)
            time.sleep(1)
            ahk.key_press('Enter')
            ahk.type("0ca")
            ahk.type("g")#general
            ahk.key_press("Tab")
            ahk.type("Debbie is retiring")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type("e") #email
            ahk.key_press("Tab")
            ahk.type("pmojoNFD")
            ahk.key_press("Tab")
            ahk.type(m+"/"+d+"/"+y)
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type("Debbie, our hygienist of 17+ years has been such an integral part of our dental family and so important to the continuity of care for our patients. We are throwing a Going-Away Party in her honor and wanted to invite you! If you are interested in coming and saying goodbye in person, please RSVP by September 15th, with the number of people attending in your party.")
            time.sleep(1)
                
            ahk.key_press("Enter")

            ahk.type("l")
            ahk.type("l")
            ahk.type("n")

    #chrome
    chrome.set_focus()


    #letter
    driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=8&cdn=1")#debretire postcard
    time.sleep(1)
    send_keys("^a^c")


    root = tk.Tk()
    root.withdraw()
    clip_text = root.clipboard_get()

    go = open('debretire.txt' ,"w",encoding="utf-8")
    go.write(clip_text)
    go.close()


    text_file = open('debretire.txt', 'r')
    text = text_file.readlines()
    text_file.close()

    del text[0:7+1]
    del text[-5:]

    new_file = open("debretire.txt", "w+")

    temp = text
    skip5=0

    alltext=""
    single=""
    lastlength = 1

    totallines = 0
    for line in temp:
        totallines += 1
        numcomma = 0
        commaindex = 0
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
            for letter in line:
                count=count+1
                if skip5>0:
                    skip5 -= 1
                elif letter == " " and count==1:
                    skip5 = 5
                elif letter == ",":
                    numcomma+=1
                    if numcomma==1:
                        commaindex=len(alltext)
                    lastlength = count
                    alltext += letter
                else:
                    alltext += letter
        #if second comma (Phd or DDS or etc removes first comma)
        if numcomma>1:
            tempt=""
            noti = commaindex
            for i in range(len(alltext)):
                if i != noti:
                    tempt = tempt + alltext[i]
            alltext = tempt
        alltext = alltext.rstrip()
        alltext+="\n"
        
    alltext = alltext[:-1]
    alltext = alltext.replace("\t","")
    alltext = alltext.replace("Bounced","")
    alltext = alltext.replace("Opt Out","")
    alltext = alltext.strip()

    for line in alltext:
        new_file.write(line)
        

    new_file.close()

    #softdent
    #app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

    num = "[0-9]+"

    NUMBER = re.compile(num)

    file1 = open("recarereallyL.txt", "r")
    lines = file1.readlines()
    file1.close()

    softdent.set_focus()


    for line in lines:
        if line != "":
                
            line = line.replace("\n","")
            ahk.key_press("Tab")
            ahk.type('f')
                
            ahk.type(line)
            time.sleep(1)
            ahk.key_press('Enter')
            ahk.type("0ca")
            ahk.type("g")#general
            ahk.key_press("Tab")
            ahk.type("Debbie is retiring")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type("l") #letter
            ahk.key_press("Tab")
            ahk.type("pmojoNFD")
            ahk.key_press("Tab")
            ahk.type(m+"/"+d+"/"+y)
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.key_press("Tab")
            ahk.type("Debbie, our hygienist of 17+ years has been such an integral part of our dental family and so important to the continuity of care for our patients. We are throwing a Going-Away Party in her honor and wanted to invite you! If you are interested in coming and saying goodbye in person, please RSVP by September 15th, with the number of people attending in your party.")
            time.sleep(1)
                
            ahk.key_press("Enter")

            ahk.type("l")
            ahk.type("l")
            ahk.type("n")

    #chrome
    chrome.set_focus()

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


