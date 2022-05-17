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



#date
today = date.today()
d = today.strftime("%d")
m = today.strftime("%m")
y = today.strftime("%Y")

#softdent
#app = Application(backend="uia").connect(title_re=".*SDWIN.EXE*")


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
    
for line in alltext:
    new_file.write(line)


new_file.close()

         
            

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
alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
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
                else:
                    alltext += letter
    
for line in alltext:
    new_file.write(line)

new_file.close()




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
alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
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
                else:
                    alltext += letter
    
for line in alltext:
    new_file.write(line)

new_file.close()



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
alltext=""
single=""
lastlength = 1

for line in temp:
    if line[0:lastlength+3] not in single:
        single += line
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
                else:
                    alltext += letter
    
for line in alltext:
    new_file.write(line)

new_file.close()



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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()



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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()



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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()


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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=33&cdn=3")#Recare: Past Due	Text Message
send_keys("^a^c")


root = tk.Tk()
root.withdraw()
clip_text = root.clipboard_get()

go = open('recarepastT.txt' ,"w",encoding="utf-8")
go.write(clip_text)
go.close()


text_file = open('recarepastT.txt', 'r')
text = text_file.readlines()
text_file.close()

del text[0:7+1]
del text[-5:]

new_file = open("recarepastT.txt", "w+")

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

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
    
for line in alltext:
    new_file.write(line)
    

new_file.close()

while(True):
    pass