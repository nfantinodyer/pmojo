import re


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
skip=False
alltext=""
single=[]
liststring=[]
lastlength = 1
totallines = 0

for line in temp:
    totallines += 1
    #if line[0:lastlength+3] in single:
    #    if re.search("@", line):
            #puts first time first no matter the order in practice mojo. Removes second occurance.
    #        for li in liststring:
    #            if li[0:lastlength+3]==line[0:lastlength+3]:
                    #since 12pm is earlier than 3pm but 12>3 so this corrects that and fixes 7am time just in case
    #               if line[next6+1:next6+3]<="12" and line[next6+1:next6+3]>="7":
    #                    if line[next6+1:next6+3] > li[next6+1:next6+3]:
    #                        single = single[0:len(single)-(len(li)-6)]
    #                        alltext = alltext[0:len(alltext)-(len(li)-6)]
                    #since 3pm is before 4pm and 3<4 
    #                elif line[next6+1:next6+3] < li[next6+1:next6+3]:
    #                    single = single[0:len(single)-(len(li)-6)]
    #                    alltext = alltext[0:len(alltext)-(len(li)-6)]    

    #skips family and gets time and date after the @ sign.
    count=0
    for letter in line:
        count=count+1
        if letter == ",":
            lastlength = count

    for i in single:
        if re.fullmatch(line[0:lastlength+3], i[0:lastlength+3]):
            skip=True
            break
    
    if not skip:
        single.append(line)
        liststring.append(line)
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
    skip=False


    alltext = alltext.rstrip()
    alltext+="\n"
    
#dont need to replace opt out and address supressed due to the next6 being the next 6 chars 00:00AM
alltext = alltext[:-1]
alltext = alltext.replace("\t","")
alltext = alltext.strip()

for line in alltext:
    new_file.write(line)

new_file.close()

#to compare if char is num
num = "[0-9]+"
NUMBER = re.compile(num)

file1 = open("practicemojo.txt", "r")
lines = file1.readlines()
file1.close()

for line in lines:
    now = False
    name = ""
    if line != "":
        for word in line:
            if re.fullmatch(NUMBER, word):
                break
            else:
                name+=word
    print(name,end=" ")
    size = 0
    for word in line:
        size+=1
        if now:
            line = line[size-1:]
            now=False
            break
        elif re.fullmatch(NUMBER, word):
            print(word,end="")
            now = True
        
    for word in line:
        if re.fullmatch("M", word):
            print("M",end="")
            break
        else:
            print(word,end="")
    print()


