#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from __future__ import print_function
from googleapiclient.discovery import build
import speech_recognition as sr
import webbrowser
import playsound
import random
import os 
import smtplib
import wikipedia
from timeit import default_timer as timer 
import datetime
import pickle
import os.path
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import time
import pyttsx3
import pytz
from selenium import webdriver
from tkinter import *
from tkinter import simpledialog
import wikipedia
from time import ctime


# In[ ]:


SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ["january", "february", "march", "april", "may", "june","july", "august", "september","october","november", "december"]
DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
DAY_EXTENTIONS = ["rd", "th", "st", "nd"]

def authenticate_google():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day, service):
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                        singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        vision_speak('No upcoming events found.')
    else:
        vision_speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("+")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0])-12) + start_time.split(":")[1]
                start_time = start_time + "pm"

            vision_speak(event["summary"] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENTIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    if month < today.month and month != -1:  
        year = year+1
    
    if month == -1 and day != -1: 
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month

    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)

    if day != -1:
        return datetime.date(month=month, day=day, year=year)

def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])

username=''
password=''
def get_id():
    def login():
        window.destroy()
    window=Tk()
    window.geometry("300x300")
    l1=Label(window,text="Student ID:",font=(12))
    l1.grid(row=0,column=0,padx=5,pady=5)
    user=StringVar()
    pas=StringVar()
    t1=Entry(window,textvariable=user,font=(12))
    t1.grid(row=0,column=1)
    l2=Label(window,text="password:",font=(12))
    l2.grid(row=1,column=0,padx=5,pady=5)
    t2=Entry(window,show='*',textvariable=pas,font=(12))
    t2.grid(row=1,column=1)
    b1=Button(window,command=login,text='Login')
    b1.grid(row=2,column=1)
    window.mainloop()
    global username,password
    password=pas.get()
    username=user.get()
    
def get_emaid():
    def login1():
        window.destroy()
    window=Tk()
    window.geometry("300x300")
    l1=Label(window,text="Email Id:",font=(12))
    l1.grid(row=0,column=0,padx=5,pady=5)
    user=StringVar()
    t1=Entry(window,textvariable=user,font=(12))
    t1.grid(row=0,column=1)
    b1=Button(window,command=login1,text='submit')
    b1.grid(row=1,column=1)
    window.mainloop() 
    return user.get()

def gmail_login():
    global username,password
    get_id()
    mail_address =username+'@charusat.edu.in'
    driver = webdriver.Chrome('chromedriver.exe')
    url = 'https://accounts.google.com/signin/v2/identifier?flowName=GlifWebSignIn&flowEntry=ServiceLogin'
    driver.get(url)
    driver.find_element_by_id("identifierId").send_keys(mail_address)
    driver.find_element_by_id("identifierNext").click()
    time.sleep(4)
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_id("passwordNext").click()

def get_my_result():  
    global username
    driver = webdriver.Chrome('chromedriver.exe')
    url = "https://charusat.edu.in:912/UniExamResult/frmUniversityResult.aspx"
    driver.get(url) 
    time.sleep(2)
    driver.find_element_by_id("ddlInst").click()
    driver.find_element_by_id("ddlInst").send_keys("CSPIT")
    time.sleep(1)
    driver.find_element_by_id("ddlDegree").click()
    driver.find_element_by_id("ddlDegree").send_keys("BTECH(IT)")
    time.sleep(1)
    driver.find_element_by_id("ddlSem").click()
    driver.find_element_by_id("ddlSem").send_keys("5")
    time.sleep(1)
    driver.find_element_by_id("ddlScheduleExam").click()
    driver.find_element_by_id("ddlScheduleExam").send_keys("DECEMBER 2020")
    time.sleep(1)
    driver.find_element_by_id("txtEnrNo").click()
    driver.find_element_by_id("txtEnrNo").send_keys(username)
    time.sleep(1)
    driver.find_element_by_id("btnSearch").click()    


# In[ ]:


import tkinter as tk
chrome_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
a=0
r = sr.Recognizer()
terms=''
def record_audio(ask = False):
    with sr.Microphone() as source:
        if ask:
            vision_speak(ask)
        audio = r.listen(source)
        voice_data = ''
        try:
            voice_data = r.recognize_google(audio)
        except sr.UnknownValueError:
            vision_speak('Sorry, I did not get that')
            global a
            a=1
        except sr.RequestError:
            vision_speak('Sorry, my speech service is down')
        return voice_data

def vision_speak(audio_string):
    en = pyttsx3.init()
    en.say(audio_string)
    print(audio_string)
    en.runAndWait()

    


def respond(voice_data):
    print(voice_data)
    def there_exists(terms):
        for term in terms:
            if term in voice_data.lower():
                return True
    
    if there_exists(['hey','hi','hello']):
        greetings = ["hey, how can I help you", "hey, what's up?", "I'm listening", "how can I help you?", "hello"]
        greet = greetings[random.randint(0,len(greetings)-1)]
        vision_speak(greet)
    
    elif 'what is your name' in voice_data.lower():
        vision_speak('This is Vision')
    
    elif 'what time is it' in voice_data.lower():
        vision_speak(ctime())
    
    elif 'google search' in voice_data.lower():
        search = record_audio('What do you want to search for?')
        url = 'https://google.com/search?q=' + search
        browser= webbrowser.get('chrome')
        browser.open_new(url)
       # webbrowser.get().open(url)
        vision_speak('Here is what I found for ' + search)
    
    elif 'find weather' in voice_data.lower():
        search = record_audio('Say a city name')
        url = 'https://google.com/search?q=' + search + 'weather'
        webbrowser.get().open(url)
        vision_speak('Here is what I found for ' + search + ' weather')    
    
    elif 'find location' in voice_data.lower():
        search = record_audio('What is the location?')
        url = 'https://www.google.com/maps/place/' + search
        webbrowser.get().open(url)
        vision_speak('Here is what I found for ' + search + ' Location')    

    elif 'make a note' in voice_data.lower():
        file = open("not10.txt", "w") 
        vision_speak("What is content?")
        tempy=record_audio()
        file.write(tempy) 
        file.close() 
        osCommandString = "notepad.exe not10.txt"
        os.system(osCommandString)
        
    elif 'charusat fees' in voice_data.lower():
        url='https://charusat.edu.in:912/FeesPaymentApp/'
        webbrowser.get().open(url)
      
    elif 'google login' in voice_data.lower():
        gmail_login()
        
    elif 'show my result' in voice_data.lower():
        get_my_result()  
        
    elif 'send email' in voice_data.lower():
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        content=record_audio('What is message')
        server.login(mail_address, password)
        server.sendmail(mail_address',get_emaid() , content)
        server.close()  
    
    elif 'search on wikipedia' in voice_data.lower():
        root = Tk() 
        ac=record_audio('What do you want to search for?') 
        root.geometry("500x500") 
        T = Text(root, height = 10, width = 100) 
        l = Label(root, text = "Output") 
        l.config(font =("Courier", 14)) 
        Fact =wikipedia.summary(ac,sentences =3 )
        # Create an Exit button. 
        b2 = Button(root, text = "Exit", command = root.destroy)  
        l.pack() 
        T.pack() 
        b2.pack() 
        T.insert(tk.END, Fact) 
        tk.mainloop()
                      
                        
    elif 'open geeksforgeeks' in voice_data.lower():
        url='https://www.geeksforgeeks.org/' 
        browser= webbrowser.get('chrome')
        browser.open_new(url)                
                        
    
    elif there_exists(["what do i have", "do i have plans", "am i busy"]):
        date = get_date(voice_data)
        if date:
            get_events(date, SERVICE)
        else:
            vision_speak("I don't understand")         
     
    else:
        global a
        if a!=1:
            vision_speak('Functionality not available')


# In[ ]:


SERVICE = authenticate_google()
print("Start")

while True:
    a=0
    vision_speak("Listening ")
    print("..........")
    voice_data = record_audio()
    respond(voice_data)

