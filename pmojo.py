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



#date
today = date.today()
d = today.strftime("%d")
m = today.strftime("%m")
y = today.strftime("%Y")


#ahk
ahk = AHK()

#chrome
driver = webdriver.Chrome() #executable_path=r"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python37_64\chromedriver.exe"
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

chrome.set_focus()
driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=21&cdn=1")#bday adult postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('bdayletter.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('bdayletter.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("bdayletter.txt", "w+")

temp = text
skip5=0
alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)


new_file.close()

#softdent

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("bdayletter.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.key_press("Tab")
    ahk.type("bday card")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()


driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=23&cdn=2")#Courtesy Reminder	Text Message
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('ReminderText.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('ReminderText.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("ReminderText.txt", "w+")

temp = text
skip5=0
next6=0
alltext=""
single=""
liststring=[]
lastlength = 1

for line in temp:
    if line[0:lastlength+3] in single:
        for li in liststring:
            if li[0:lastlength+3]==line[0:lastlength+3]:
                if line[next6+1:next6+3]<="12" and line[next6+1:next6+3]>="7":
                    if line[next6+1:next6+3] > li[next6+1:next6+3]:
                        single = single[0:len(single)-len(li)]
                        alltext = alltext[0:len(alltext)-(len(li)-6)]
                elif line[next6+1:next6+3] < li[next6+1:next6+3]:
                    single = single[0:len(single)-len(li)]
                    alltext = alltext[0:len(alltext)-(len(li)-6)]
               
    
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
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("ReminderText.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    for word in line.split():
        if re.fullmatch(NUMBER, word):
            break
        else:
            ahk.type(word)
    ahk.key_press("Enter")
    ahk.type("0ca")
    ahk.type("cc")
    ahk.key_press("Tab")
    ahk.type("Reminder for ")
    for word in line.split():
        if re.fullmatch(NUMBER, word):
            ahk.type(word)
        elif re.fullmatch("M", word):
            ahk.type("M")
            stop=True
            break

    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("t") #text
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()


driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=23&cdn=1")#Courtesy Reminder	Email
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('ReminderE.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('ReminderE.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("ReminderE.txt", "w+")

temp = text
skip5=0
next6=0
alltext=""
single=""
liststring=[]
lastlength = 1


for line in temp:
    if line[0:lastlength+3] in single:
        for li in liststring:
            if li[0:lastlength+3]==line[0:lastlength+3]:
                if line[next6+1:next6+3]<="12" and line[next6+1:next6+3]>="7":
                    if line[next6+1:next6+3] > li[next6+1:next6+3]:
                        single = single[0:len(single)-len(li)]
                        alltext = alltext[0:len(alltext)-(len(li)-6)]
                elif line[next6+1:next6+3] < li[next6+1:next6+3]:
                    single = single[0:len(single)-len(li)]
                    alltext = alltext[0:len(alltext)-(len(li)-6)]
               
    
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
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("ReminderE.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    for word in line.split():
        if re.fullmatch(NUMBER, word):
            break
        else:
            ahk.type(word)
    ahk.key_press("Enter")
    ahk.type("0ca")
    ahk.type("cc")
    ahk.key_press("Tab")
    ahk.type("Reminder for ")
    for word in line.split():
        if re.fullmatch(NUMBER, word):
            ahk.type(word)
        elif re.fullmatch("M", word):
            ahk.type("M")
            stop=True
            break

    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("e") #email
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")
    
    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=130&cdn=2")#Courtesy Reminder: Unconfirmed	Text Message
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('UnconfirmedText.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('UnconfirmedText.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("UnconfirmedText.txt", "w+")

temp = text
skip5=0
next6=0
alltext=""
single=""
liststring=[]
lastlength = 1


for line in temp:
    if line[0:lastlength+3] in single:
        for li in liststring:
            if li[0:lastlength+3]==line[0:lastlength+3]:
                if line[next6+1:next6+3]<="12" and line[next6+1:next6+3]>="7":
                    if line[next6+1:next6+3] > li[next6+1:next6+3]:
                        single = single[0:len(single)-len(li)]
                        alltext = alltext[0:len(alltext)-(len(li)-6)]
                elif line[next6+1:next6+3] < li[next6+1:next6+3]:
                    single = single[0:len(single)-len(li)]
                    alltext = alltext[0:len(alltext)-(len(li)-6)]
               
    
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
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("UnconfirmedText.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    for word in line.split():
        if re.fullmatch(NUMBER, word):
            break
        else:
            ahk.type(word)
    ahk.key_press("Enter")
    ahk.type("0ca")
    ahk.type("cc")
    ahk.key_press("Tab")
    ahk.type("Reminder for ")
    for word in line.split():
        if re.fullmatch(NUMBER, word):
            ahk.type(word)
        elif re.fullmatch("M", word):
            ahk.type("M")
            stop=True
            break

    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("t") #text
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")
    
    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=36&cdn=1")#e-Birthday Adult	Email
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('ebday.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('ebday.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("ebday.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter

alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("ebday.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.key_press("Tab")
    ahk.type("bday card")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("e") #email
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=22&cdn=1")#Bday Child Postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('bdaykid.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('bdaykid.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("bdaykid.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter

alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("bdaykid.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.key_press("Tab")
    ahk.type("bday card")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()


driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=35&cdn=1")#Happy Anniversary	Postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('anniversary.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('anniversary.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("anniversary.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("anniversary.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.key_press("Tab")
    ahk.type("anniversary card")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()


driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=30&cdn=2")#Reactivate: 1 year ago	Email
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('reactivateE.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('reactivateE.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("reactivateE.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("reactivateE.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.key_press("Tab")
    ahk.type("Reactivate: 1 year ago")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("e") #email
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=30&cdn=1")#Reactivate: 1 year ago	Postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('reactivateL.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('reactivateL.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("reactivateL.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("reactivateL.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.key_press("Tab")
    ahk.type("Reactivate: 1 year ago")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=1&cdn=2")#Recare: Due	Email
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recaredueE.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recaredueE.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recaredueE.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("recaredueE.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("e") #email
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=1&cdn=3")#Recare: Due	Text Message
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recaredueT.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recaredueT.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recaredueT.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("recaredueT.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("t") #txt
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=1&cdn=1")#Recare: Due    Postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recaredueL.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recaredueL.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recaredueL.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter

alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("recaredueL.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=33&cdn=1")#Recare: Past Due	Postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recarepastL.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recarepastL.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recarepastL.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter

alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("recarepastL.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Past Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=33&cdn=2")#Recare: Past Due	Email
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recarepastE.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recarepastE.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recarepastE.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("recarepastE.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Past Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("e") #email
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=8&cdn=2")#Recare: Really Past Due	Email
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recarereallyE.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recarereallyE.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recarereallyE.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter

alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

for line in alltext:
    new_file.write(line)
    

new_file.close()

#softdent
#app = Application().start("C:\Program Files (x86)\SoftDent\SoftDent.exe")

num = "[0-9]+"

NUMBER = re.compile(num)

file1 = open("recarereallyE.txt", "r")
lines = file1.readlines()
file1.close()

softdent.set_focus()

for line in lines:
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Really Past Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("e") #email
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=8&cdn=1")#Recare: Really Past Due	Postcard
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recarereallyL.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recarereallyL.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recarereallyL.txt", "w+")

temp = text
skip5=0

alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
        if re.search("Family", line):
            continue
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
            else:
                alltext += letter
    
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.replace(" ","")

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
    line = line.replace("\n","")
    ahk.type('f')
    ahk.type(line)
    ahk.key_press('Enter')
    ahk.type("0ca")
    ahk.type("r")#recare
    ahk.key_press("Tab")
    ahk.type("Recare: Really Past Due")
    ahk.key_press("Tab")
    ahk.key_press("Tab")
    ahk.type("l") #letter
    ahk.key_press("Tab")
    ahk.type("pmojoNFD")
    ahk.key_press("Tab")
    ahk.key_press("Enter")

    ahk.type("l")
    ahk.type("l")
    ahk.type("n")

    

#chrome
chrome.set_focus()

while(True):
    pass