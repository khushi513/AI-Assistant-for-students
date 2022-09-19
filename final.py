import datetime
import json
import os
import random
import time
import webbrowser
import pygame

import pyautogui
import pyjokes
import pyttsx3
import pywhatkit
import requests
import speech_recognition as sr
import wikipedia
from bs4 import BeautifulSoup
from playsound import playsound

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
# print(voices[1].id)
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon!")   

    else:
        speak("Good Evening!")  

    speak("I am your personal assistant. Please tell me how may I help you")       

def takeCommand():
    #It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 2
        audio = r.listen(source)

    try:
        print("Recognizing...")    
        query = r.recognize_google(audio, language='en-in')
        print("User said: {query}\n")

    except Exception :    
        print("Say that again please...")  
        return "None"
    return query

def NewsFromBBC():
    
    # BBC news api
    # following query parameters are used
    # source, sortBy and apiKey
    query_params = {
        "source": "bbc-news",
        "sortBy": "top",
        "apiKey": "d83eb8226fd6484bb854940991a09614"
    }
    main_url = " https://newsapi.org/v1/articles"

    res = requests.get(main_url, params=query_params)
    open_bbc_page = res.json()

    article = open_bbc_page["articles"]

    results = []

    for ar in article:
        results.append(ar["title"])

    for i in range(len(results)):

        print(i + 1, results[i])

    from win32com.client import Dispatch
    speak = Dispatch("SAPI.Spvoice")
    speak.Speak(results)

def weather_data(query):
    res = requests.get('http://api.openweathermap.org/data/2.5/weather?' +
                       query+'&APPID=42b556180bf745dc9ce8b7ab71bcac36&units=metric')
    return res.json()

def print_weather(result, city):
    print("{}'s temperature: {}°C ".format(city, result['main']['temp']))
    speak("{}'s temperature: {}°C ".format(city, result['main']['temp']))
    print("Wind speed: {} m/s".format(result['wind']['speed']))
    speak("Wind speed: {} m/s".format(result['wind']['speed']))
    print("Description: {}".format(result['weather'][0]['description']))
    speak("Description: {}".format(result['weather'][0]['description']))
    print("Weather: {}".format(result['weather'][0]['main']))
    speak("Weather: {}".format(result['weather'][0]['main']))

def TakeSS():
    ss = pyautogui.screenshot()
    ss.save("C:\\Users\\HP\\OneDrive - Vivekananda Institute of Professional Studies\\Desktop\\gp\\ss\\ss.jpg")
    speak("I have taken the screenshot")
    ss.show()
    

if __name__ == "__main__":
    wishMe()
    while True:
    # if 1:
        query = takeCommand().lower()

        # Logic for executing tasks based on query
        if 'what can you do' in query:
            speak("""I can do a lot of things like,set alarms and reminders, join classes for you
         ,search on wikipedia,calculate,open applications,play music,show weather,open maps,send message,etc""")

        elif 'hi' in query:
            speak("Hi! What can i do for you?")

        elif 'how are you' in query:
            r = ['good', 'fine', 'great']
            response = random.choice(r)
            print(f"I am {response}")
            t2s(f"I am {response}")
            

        elif "wikipedia" in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2) 
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            print("Opening Youtube") 
            speak("Opening Youtube") 
            webbrowser.open("https://www.youtube.com/")

        elif 'open google chrome' in query:
            print("Opening Google Chrome") 
            speak("Opening Google Chrome") 
            webbrowser.open("https://www.google.com/")

        elif 'stack overflow' in query:
            print("Opening stackoverflow") 
            speak("Opening stackoverflow") 
            webbrowser.open("https://stackoverflow.com/")

        elif "word" in query: 
            print("Opening Microsoft Word") 
            speak("Opening Microsoft Word") 
            os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")

        elif "excel" in query: 
            print("Opening Microsoft Excel") 
            speak("Opening Microsoft Excel") 
            os.startfile("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
        
        elif 'play music' in query:
            print("Playing music") 
            speak("Playing music") 
            music = 'C:\\Users\\HP\\OneDrive - Vivekananda Institute of Professional Studies\\Desktop\\gp\\music'
            songs = os.listdir(music)
            print(songs)    
            os.startfile(os.path.join(music, songs[0]))
            time.sleep(46)

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M")    
            speak(f"Sir, the time is {strTime}")
            print(f"Sir, the time is {strTime}")

        elif 'weather' in query:
            speak("Enter name of the city: ")
            print("Enter name of the city: ")
            city = input()
            print()
            try:
                query = 'q='+city
                w_data = weather_data(query)
                print_weather(w_data, city)
                print()
            except:
                print('City name not found...')

        elif "temperature" in query:
             search = "temperature in delhi"
             url = f"https://www.google.com/search?q={search}"
             r = requests.get(url)
             data = BeautifulSoup(r.text,"html.parser")
             temp = data.find("div",class_="BNeawe").text
             print(f"current {search} is {temp}")
             speak(f"current {search} is {temp}")

        elif 'open code' in query:
            print("Opening vs code")
            speak("Opening vs code")
            codepath = "C:\\Program Files\\Microsoft VS Code\\Code.exe"
            os.startfile(codepath)

        elif 'joke' in query:
            joke=pyjokes.get_joke()
            engine.setProperty("rate",178)
            print(joke)
            speak(joke)

        elif 'toss a coin' in query or 'flip a coin' in query:
            coin_flip = ['heads','tails']
            toss = random.choice(coin_flip)
            print ("Its "+ toss)
            speak ("Its "+ toss)

        elif 'rock paper scissors' in query or 'rock paper scissor' in query:
            speak("choose among rock paper or scissor")
            print("I choose: ")
            pmove=input("")
            cmove=random.choice(['rock', 'paper', 'scissor'])
            speak("The computer chose " + cmove)
            speak("You chose " + pmove)

            if pmove==cmove:
                speak("the match is draw")
            elif pmove== "rock" and cmove== "scissor":
                speak("You win")
            elif pmove== "rock" and cmove== "paper":
                speak("Computer wins")
            elif pmove== "paper" and cmove== "rock":
                speak("You win")
            elif pmove== "paper" and cmove== "scissor":
                speak("Computer wins")
            elif pmove== "scissor" and cmove== "paper":
                speak("You win")
            elif pmove== "scissor" and cmove== "rock":
                speak("Computer wins")

        elif 'classes' in query or 'schedule' in query:
            print('Heres your schedule for todays classes')
            speak('Heres your schedule for todays classes')
            webbrowser.open('https://teams.microsoft.com/_?culture=en-in&country=IN&lm=deeplink&lmsrc=homePageWeb&cmpid=WebSignIn#/calendarv2')

        elif "write a note" in query:
            print("What should i write, sir")
            speak("What should i write, sir")
            note = takeCommand()
            file = open('notes.txt', 'w')
            print("Sir, Should i include date and time")
            speak("Sir, Should i include date and time")
            query = takeCommand()
            if 'yes' in query or 'sure' in query:
                strTime = datetime.datetime.now().strftime("%H:%M")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show notes" in query or 'open notes' in query:
            print("Showing Notes")
            speak("Showing Notes")
            file = open("notes.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "reminder" in query or "remind me" in query:
            print("What shall I remind you about?")
            speak("What shall I remind you about?")
            text = takeCommand()
            print("In how many minutes?")
            speak("In how many minutes?")
            local_time = int(takeCommand())
            local_time = local_time * 60
            time.sleep(local_time)
            print(text)
            speak(text)
            freq=100
            dur=50
            for i in range(0, 5):    
                winsound.Beep(freq,dur) 
                freq+= 100
                dur+= 50

        elif 'set alarm' in query:
            speak("Enter time: ")
            time= input("Enter time: ")
            while True:
                Time_Ac=datetime.datetime.now()
                now= Time_Ac.strftime("%H:%M")
                if now==time:
                    i=1
                    while (i<6):
                        speak("Wake up!")
                        print("Wake up!")
                        playsound("alarm.wav")
                        i+=1
                    break
            speak("Alarm closed.")
            print("Alarm closed.")

        elif 'mail' in query or 'check my mail' in query:
            search_term = query.split("for")[-1]
            url="https://mail.google.com/mail/u/0/#inbox"
            webbrowser.get().open(url)
            speak("here you can check your gmail")

        elif 'send message' in query or 'whatsapp' in query:
            print("What message should i send?")
            speak("What message should i send?")
            message = takeCommand()
            speak("Enter time")
            time = input("Enter time in HH:MM\n-")
            hour, minute = [int(i) for i in time.split(":")]
            pywhatkit.sendwhatmsg('+919899375111',message,hour,minute)
            print("Message sent!")
            speak("Message sent!")

        elif 'news' in query:
            NewsFromBBC()

        elif "screenshot" in query:
            TakeSS()

        elif 'calculator' in query:
            def add(num1, num2):
                return num1 + num2
            
            # Function to subtract two numbers 
            def subtract(num1, num2):
                return num1 - num2
            
            # Function to multiply two numbers
            def multiply(num1, num2):
                return num1 * num2
            
            # Function to divide two numbers
            def divide(num1, num2):
                return num1 / num2
            
            print("Please select operation -\n" \
                    "1. Add\n" \
                    "2. Subtract\n" \
                    "3. Multiply\n" \
                    "4. Divide\n")

            speak("Select operations")
            select = int(input("Select operations from 1, 2, 3, 4 : "))
            
            number_1 = int(input("Enter first number: "))
            number_2 = int(input("Enter second number: "))
            
            if select == 1:
                print (number_1, "+", number_2, "=",add(number_1, number_2))
            elif select == 2:
                print (number_1, "-", number_2, "=",subtract(number_1, number_2))

            elif select == 3:
                print (number_1, "*", number_2, "=",multiply(number_1, number_2))

            elif select == 4:
                print (number_1, "/", number_2, "=",divide(number_1, number_2))

            else:
                print("Invalid input")

        elif 'thank you' in query:
            speak("My pleasure. Its good to bee acknowledged.")
            print("My pleasure. Its good to bee acknowledged.")

        elif 'exit' in query:
            speak("Thanks for using my services. Have a good day! ")
            exit()
        
        else:
            print("Sorry, I could not understand you.")
            speak("Sorry, I could not understand you.")