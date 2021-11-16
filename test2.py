import speech_recognition as sr
import playsound
import urllib.request
import re
import webbrowser as wbs
import xlrd as excel
import datetime
import calendar as call
import pyttsx3
import requests as req

weekdays = ['MONDAY','TUESDAY','WEDNESDAY','THURSDAY','FRIDAY','SATURDAY','SUNDAY']
definedtime = ['08:00:AM','8:50:AM','9:40:AM','10:30:AM','11:20:AM','12:10:PM','13:00:PM','13:50:PM','14:40:PM','15:30:PM']

engine = pyttsx3.init()
engine.setProperty('voice', engine.getProperty('voices')[1].id)
engine.setProperty('rate', 125)

def speak(text):
    engine.say(text)
    engine.runAndWait()    

def get_audio(text):
    r = sr.Recognizer()
    mic = sr.Microphone()
    with mic as source:
        r.adjust_for_ambient_noise(source)
        print("Speak Now")
        if text=='0':
            print("",end="")
        else:
            speak(text)
        a=r.listen(source)
        j=""
        try:
            j=r.recognize_google(a)
            j=j.lower()
        except Exception as e:
            print("",end="")
    return j
    
def make_a_note():
	check = open(r"C:\Users\ASUS\OneDrive\Pictures\Screenshots\OneDrive\Desktop\PYTHON\check.txt")
	number=check.read()
	check.close()
	k=number
	rewrite = open(r"C:\Users\ASUS\OneDrive\Pictures\Screenshots\OneDrive\Desktop\PYTHON\check.txt","w")
	rewrite.write(str(int(k)+1))
	rewrite.close()
	file = open(str(datetime.date.today())+number+".txt","w")
	file.write(get_audio("What do you want me to write"))
	file.close()

def timetable():
    wb = excel.open_workbook(r"C:\Users\ASUS\OneDrive\Pictures\Screenshots\OneDrive\Desktop\TIME TABLE.xlsx")
    sheet = wb.sheet_by_index(0)
    c=str(datetime.date.today())
    c=c.replace('-',' ')
    born = datetime.datetime.strptime(c, '%Y %m %d').weekday()
    if(datetime.datetime.now().hour>16):
        day=call.day_name[born+1]
    else:
        day=call.day_name[born]
    day=day.upper()
    sheet.cell_value(0,0)
    if day in ['SATURDAY','SUNDAY']:
        playsound.playsound(r"C:\Users\ASUS\OneDrive\Pictures\Screenshots\OneDrive\Desktop\PYTHON\final.wav")
    else:
        for i in range(sheet.ncols):
            if day in sheet.cell_value(0,i):
                speak("The time table for "+day+" is as follows")
                for j in range(1,sheet.nrows,2):
                    if sheet.cell_value(j,i)=="":
                        continue
                    if sheet.cell_value(j,i)!="END":
                        print(str(sheet.cell_value(j,0))+"\t\t"+str(sheet.cell_value(j,i)))
                        k=sheet.cell_value(j,i)
                        m=sheet.cell_value(j,0)        
                        if 'L' in k:
                            k=k.replace('L','')
                            m=m.replace(':AM','')        
                            m=m.replace(':PM','')        
                            speak("You have a lecture "+str(k)+"at "+str(m))        
                        elif 'P' in k:        
                            k=k.replace('P','')
                            m=m.replace(':AM','')
                            m=m.replace(':PM','')
                            speak("You have a Practical "+str(k)+"at "+str(m))
                        elif 'T' in k:
                            k=k.replace('T','')
                            m=m.replace(':AM','')
                            m=m.replace(':PM','')
                            speak("You have a Tutorial "+str(k)+"at "+str(m))            
                    elif sheet.cell_value(j,i)=="END":
                        print(str(sheet.cell_value(j,0))+"\t\t"+str(sheet.cell_value(j,i)))
                        break
                    else:
                        continue
def loginzoom(x):
    wb = excel.open_workbook(r"C:\Users\ASUS\OneDrive\Pictures\Screenshots\OneDrive\Desktop\TIME TABLE.xlsx")
    sheet = wb.sheet_by_index(1)
    c=str(datetime.date.today())
    c=c.replace('-',' ')
    born = datetime.datetime.strptime(c, '%Y %m %d').weekday()
    day=call.day_name[born]
    day=day.upper()
    print(x)
    sheet.cell_value(0,0)
    if day in ['SATURDAY','SUNDAY']:
        playsound.playsound(r"C:\Users\ASUS\OneDrive\Pictures\Screenshots\OneDrive\Desktop\PYTHON\final.wav")
    else:
        for i in range(sheet.nrows):
            if x in sheet.cell_value(i,0):
                for j in range(sheet.ncols):
                    if day in sheet.cell_value(0,j):
                        link = sheet.cell_value(i+1,j)
                        print(link)
                        wbs.open(link)
                        speak("Loging in to zoom")
                        break

def song(text):
    x=txt.replace(" ","+")
    html = urllib.request.urlopen("https://www.youtube.com/results?search_query="+x)
    ids = re.findall(r"watch\?v=(\S{11})",html.read().decode())
    link="https://www.youtube.com/watch?v="+ids[0]
    return (link)


'''speak("Playing Ganesh Aarti")
wbs.open('https://www.youtube.com/watch?v=UWJSHE-OGc4')'''
while 1:
    k=get_audio('0')
    if k=="":
        print(k,end="")
    else:
        print(k)
    if "wake up" in k:
        h=get_audio("How can I help you")
        print(h)
        if "play" and "song" in h:
            txt=get_audio("What do you want to listen")
            wbs.open_new(song(txt))
        elif "time table" in h:
            speak("Opening Excel file wait")
            timetable()
        elif "open" and "mail" in h:
            clip=get_audio("Which inbox would you like me to access")
            if "thapar" in clip:
                wbs.open("https://mail.google.com/mail/u/0/#inbox")
            else:
                wbs.open("https://mail.google.com/mail/u/1/#inbox")
        elif "open" and "lms" in h:
            wbs.open("https://ada-lms.thapar.edu/moodle/my/")
        elif "make" and "note" in h:
            make_a_note()
        elif "open" and "google" in h:
            wbs.open("https://www.google.com/")
        elif "zoom" in h:
            print(definedtime)
            x = int(input())
            loginzoom(definedtime[x])
        else:
            continue
    elif "stop" in k:
        break
    elif 'make' and 'note' in k:
        make_a_note()
    elif "what" and "name" in k:
        speak("My name is nothing")
    else:
        continue
