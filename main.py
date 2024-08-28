import pyttsx3
import requests
import speech_recognition as sr
import keyboard
import os
import subprocess as sp
import imdb
import wolframalpha
import pyautogui
import webbrowser
import time

from datetime import datetime
from decouple import config
from random import choice
from conv import random_text
from online import find_my_ip, search_on_google, search_on_wikipedia, youtube, send_email, get_news, weather_forecast

engine = pyttsx3.init('sapi5')
engine.setProperty('volume',1.5)
engine.setProperty('rate',225)
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)

USER = config('USER')
HOSTNAME = config('BOT')


def speak(text):
    engine.say(text)
    engine.runAndWait()


def greet_me():
    hour = datetime.now().hour
    if(hour>=6) and (hour < 12):
        speak(f"Good morning {USER}")
    elif (hour >= 12) and (hour <=16 ):
        speak(f"Good afternoon {USER}")
    elif (hour >= 16) and (hour < 19):
        speak(f"Good evening {USER}")
    speak(f"Hi I am {HOSTNAME}. How may I assist you? {USER}")


listening =  False


def start_listening():
    global listening
    listening = True
    print("started listening ")


def pause_listening():
    global listening
    listening = False
    print("stopped listening")


keyboard.add_hotkey('ctrl+alt+k',start_listening)
keyboard.add_hotkey('ctrl+alt+p',pause_listening)


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing....")
        queri = r.recognize_google(audio,language='en-in')
        print(queri)
        if not 'stop' in queri or 'exit' in queri:
            speak(choice(random_text))
        else:
            hour  =datetime.now().hour
            if hour >= 21 and hour < 6:
                speak("Good night sir, take care!")
            else:
                speak("Have a good day sir!")
            exit()

    except Exception:
        speak("Sorry I couldn't understand. Can you please repeat that?")
        queri = 'None'
    return queri


if __name__ == '__main__':
    greet_me()
    while True:
        if listening:
            query = take_command().lower()
            if "how are you" in query:
                speak("I am absolutely fine sir. What about you")
            elif "open command prompt" in query:
                speak("Opening command prompt")
                os.system('start cmd')


            elif "open camera" in query:
                speak("Opening camera sir")
                sp.run('start microsoft.windows.camera:',shell=True)

            elif "open notepad" in query:
                speak("Opening Notepad for you sir")
                notepad_path = "C:\Windows\System32\notepad.exe"
                os.startfile(notepad_path)

            elif "open discord" in query:
                speak("Opening Discord for you sir")
                discord_path = "C:\\Users\\harsh\\AppData\\Local\\Discord\\app-1.0.9157"
                os.startfile(discord_path)

            elif "open excel" in query:
                speak("Opening Excel for you sir")
                excel_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
                os.startfile(excel_path)

            elif "open powerpoint" in query:
                speak("Opening Powerpoint for you sir")
                powerpoint_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
                os.startfile(powerpoint_path)

            elif "open onenote" in query:
                speak("Opening OneNote for you sir")
                onenote_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\ONENOTE.EXE"
                os.startfile(onenote_path)

            elif "open outlook" in query:
                speak("Opening Outlook for you sir")
                outlook_path = "C:\\Program Files\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"
                os.startfile(outlook_path)

            elif "open openshot video editor" in query:
                speak("Opening Outlook for you sir")
                openshotvideoeditor_path = "C:\\Program Files\\OpenShot Video Editor\\openshot-qt.exe"
                os.startfile(openshotvideoeditor_path)

            elif "open word" in query:
                speak("Opening Word for you sir")
                word_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
                os.startfile(word_path)

            elif "open microsoft edge" in query:
                speak("Opening Microsoft Edge for you sir")
                microsoftedge_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
                os.startfile(microsoftedge_path)

            elif "open vmware horizon view client" in query:
                speak("Opening VMware Horizon View Client for you sir")
                vmwarehorizonclient_path = "C:\\Program Files\\VMware\\VMware Horizon View Client\\vmware-view.exe"
                os.startfile(vmwarehorizonclient_path)

            elif "open google chrome" in query:
                speak("Opening Google Chrome for you sir")
                googlechrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                os.startfile(googlechrome_path)

            elif "ip address" in query:
                ip_address = find_my_ip()
                speak(
                    f"your ip address is {ip_address}"
                )
                print(f"your ip address is {ip_address}")

            elif "open youtube" in query:
                speak("What do you want to play on youtube sir?")
                video = take_command().lower()
                youtube(video)

            elif "open google" in query:
                speak("What do you want to search on google?")
                query = take_command().lower()
                search_on_google(query)

            elif "wikipedia" in query:
                speak("What do you want to search on wikipedia sir?")
                search = take_command().lower()
                results = search_on_wikipedia(search)
                speak(f"According to wikipedia,{results}")
                speak("I am printing in on terminal")
                print(results)


            elif "send an email" in query:
                speak("On what email address do you want to send sir?. Please enter in the terminal")
                receiver_add = input("Email address:")
                speak("What should be the subject sir?")
                subject = take_command().capitalize()
                speak("What is the message ?")
                message = take_command().capitalize()
                if send_email(receiver_add, subject, message):
                    speak("I have sent the email sir")
                    print("I have sent the email sir")

                else:
                    speak("something went wrong Please check the error log")

            elif "give me news" in query:
                speak(f"I am reading out the latest headlines of today,sir")
                speak(get_news())
                speak("I am printing it on screen sir")
                print(*get_news(),sep='\n')



            elif 'weather' in query:
                ip_address = find_my_ip()
                speak("tell me the name of your city")
                city = input("Enter name of your city: ")
                weather, temp, feels_like = weather_forecast(city)
                if weather:
                    speak(f"The current temperature is {temp}, but it feels like {feels_like}")
                    speak(f"Also, the weather report talks about {weather}")
                    print(f"Description: {weather}\nTemperature: {temp}\nFeels like: {feels_like}")
                else:
                    speak("Sorry, I couldn't fetch the weather information.")


            elif "movie" in query:
                movies_db = imdb.IMDb()
                speak("Please tell me the movie name:")
                text = take_command()
                movies = movies_db.search_movie(text)
                speak("searching for" + text)
                speak("I found these")
                for movie in movies:
                    title = movie["title"]
                    year = movie["year"]
                    speak(f"{title}-{year}")
                    info = movie.getID()
                    movie_info = movies_db.get_movie(info)
                    rating = movie_info["rating"]
                    cast = movie_info["cast"]
                    actor = cast[0:5]
                    plot = movie_info.get('plot outline', 'plot summary not available')
                    speak(f"{title} was released in {year} has imdb ratings of {rating}.It has a cast of {actor}. "
                          f"The plot summary of movie is {plot}")

                    print(f"{title} was released in {year} has imdb ratings of {rating}.\n It has a cast of {actor}. \n"
                          f"The plot summary of movie is {plot}")


            elif "calculate" in query:
                app_id = "4P5AK4-385W927T7L"
                client = wolframalpha.Client(app_id)
                ind = query.lower().split().index("calculate")
                text = query.split()[ind + 1:]
                result = client.query(" ".join(text))
                try:
                    ans = next(result.results).text
                    speak("The answer is " + ans)
                    print("The answer is " + ans)
                except StopIteration:
                    speak("I couldn't find that . Please try again")


            elif 'what is' in query or 'who is' in query or 'which is' in query:
                app_id = "4P5AK4-5W8AVXAYGY"
                client = wolframalpha.Client(app_id)
                try:

                    ind = query.lower().index('what is') if 'what is' in query.lower() else \
                        query.lower().index('who is') if 'who is' in query.lower() else \
                            query.lower().index('which is') if 'which is' in query.lower() else None

                    if ind is not None:
                        text = query.split()[ind + 2:]
                        res = client.query(" ".join(text))
                        ans = next(res.results).text
                        speak("The answer is " + ans)
                        print("The answer is " + ans)
                    else:
                        speak("I couldn't find that. Please try again.")
                except StopIteration:
                    speak("I couldn't find that. Please try again.")

            elif 'subscribe' in query:
                speak(
                    "Everyone who are watching this video, Please subscribe to Carbonium. "
                    "I will show you how to do this")
                speak("Firstly Go to youtube")
                webbrowser.open("https://www.youtube.com/")
                speak("click on the search bar")
                pyautogui.moveTo(806, 125, 1)
                pyautogui.click(x=806, y=125, clicks=1, interval=0, button='left')
                speak("Carbonium13")
                pyautogui.typewrite("Carbonium13", 0.1)
                time.sleep(1)
                speak("press enter")
                pyautogui.press('enter')
                pyautogui.moveTo(971, 314, 1)
                speak("Here you will see our channel")
                pyautogui.moveTo(1688, 314, 1)
                speak("click here to subscribe our channel")
                pyautogui.click(x=1688, y=314, clicks=1, interval=0, button='left')
                speak("And also Don't forget to press the bell icon")
                pyautogui.moveTo(1750, 314, 1)
                pyautogui.click(x=1750, y=314, clicks=1, interval=0, button='left')
                speak("turn on all notifications")
                pyautogui.click(x=1750, y=320, clicks=1, interval=0, button='left')