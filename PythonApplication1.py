from selenium import webdriver
import chromedriver_binary
from selenium.webdriver.common.keys import Keys
from datetime import date
from pywinauto.application import Application
import pywinauto.mouse as mouse
import pywinauto.keyboard as keyboard
from pywinauto.keyboard import send_keys
import os
import os.path
import shutil
import tkinter as tk

def erase(file_name: str, start_key: str, stop_key: str):
    """
    This function will delete all line from the givin start_key
    until the stop_key. (include: start_key) (exclude: stop_key)
    """
    try: 
        # read the file lines
        with open(file_name, 'r+') as fr: 
            lines = fr.readlines()
        # write the file lines except the start_key until the stop_key
        with open(file_name, 'w+') as fw:
            # delete variable to control deletion
            delete = False
            # iterate over the file lines
            for line in lines:
                # check if the line is a start_key
                # set delete to True (start deleting)
                    if line.strip(' ') == start_key:
                        delete = True
                    # check if the line is a stop_key
                    # set delete to False (stop deleting)
                    elif line.strip(' ') == stop_key:
                        delete = False
                    # write the line back based on delete value
                    # if the delete setten to True this will
                    # not be executed (the line will be skipped)
                    if not delete: 
                        fw.write(line)
    except RuntimeError as ex: 
        print(f"erase error:\n\t{ex}")
    

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
elem = driver.find_element_by_name("loginId")
elem.clear()
elem.send_keys("oakmeadowdental")
elem = driver.find_element_by_name("password")
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
for line in text:
    new_file.write(line)

new_file.close()

text_file = open('bdayletter.txt', 'r')
text = text_file.read()
text_file.close()
words = text.split()
words = [word.strip('.!;()[]') for word in words]
new_file = open("bdayletter.txt", "w+")
for words in text:
    new_file.write(words)

new_file.close()


#####################################################################################################                
            

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
for line in text:
    new_file.write(line)

new_file.close()

text_file = open('ReminderText.txt', 'r')
text = text_file.read()
text_file.close()
words = text.split()
words = [word.strip('.!;()[]') for word in words]
words = [word.replace('Yes','') for word in words]
new_file = open("ReminderText.txt", "w+")
for words in text:
    new_file.write(words)

new_file.close()

erase('ReminderText.txt',  "Address","\n")


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
for line in text:
    new_file.write(line)

new_file.close()

text_file = open('ReminderE.txt', 'r')
text = text_file.read()
text_file.close()
words = text.split()
words = [word.strip('.!;()[]') for word in words]
words = [word.replace('Yes','') for word in words]
new_file = open("ReminderE.txt", "w+")
for words in text:
    new_file.write(words)

new_file.close()

erase('ReminderE.txt',  "Address","\n")

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
for line in text:
    new_file.write(line)

new_file.close()

text_file = open('UnconfirmedText.txt', 'r')
text = text_file.read()
text_file.close()
words = text.split()
words = [word.strip('.!;()[]') for word in words]
new_file = open("UnconfirmedText.txt", "w+")
for words in text:
    new_file.write(words)

new_file.close()

erase('UnconfirmedText.txt', "Address","\n")

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
for line in text:
    new_file.write(line)

new_file.close()

text_file = open('ebday.txt', 'r')
text = text_file.read()
text_file.close()
words = text.split()
words = [word.strip('.!;()[]') for word in words]
new_file = open("ebday.txt", "w+")
for words in text:
    new_file.write(words)

new_file.close()

erase('ebday.txt', "Address","\n")

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
for line in text:
    new_file.write(line)

new_file.close()

text_file = open('anniversary.txt', 'r')
text = text_file.read()
text_file.close()
words = text.split()
words = [word.strip('.!;()[]') for word in words]
new_file = open("anniversary.txt", "w+")
for words in text:
    new_file.write(words)

new_file.close()

erase('anniversary.txt', "Address","\n")

#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=30&cdn=2")#Reactivate: 1 year ago	Email


#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=30&cdn=1")#Reactivate: 1 year ago	Postcard


#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=1&cdn=2")#Recare: Due	Email


#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=1&cdn=3")#Recare: Due	Text Message


#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=33&cdn=2")#Recare: Past Due	Email


#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=8&cdn=2")#Recare: Really Past Due	Email


#driver.get("https://app.practicemojo.com/cgi-bin/WebObjects/PracticeMojo.woa/wa/gotoActivityDetail?td="+m+"%2F"+d+"%2F"+y+"&cdi=8&cdn=1")#Recare: Really Past Due	Postcard


while(True):
    pass
