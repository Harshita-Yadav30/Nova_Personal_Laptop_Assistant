#C:\\Users\\harsh\\AppData\\Local\\Programs\\Python\\Python310\\python.exe .\nova.py

import speech_recognition as sr
import pyttsx3
import datetime, random, time
import os
import subprocess

r = sr.Recognizer()
engine = pyttsx3.init()

engine.setProperty('rate', 150)
voice = engine.getProperty('voices')[1]
engine.setProperty('voice', voice.id)

def greet():
    #Greeting the user at start of the execution
    hour = datetime.datetime.now().hour
    if hour < 12:
        engine.say("Good Morning!")
    elif hour < 16:
        engine.say("Good Afternoon!")
    else:
        engine.say("Good Evening!")
    
    engine.say("I'm Nova and I'll be assisting you in your tasks.")
    engine.say("Please give your command and I'll try to do them for you")
    engine.runAndWait()

def listen():
    #Listening to user here
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.05)
        print("Listening...")

        while True:
            try:
                audio = r.listen(source)
                text = r.recognize_google(audio)
                # engine.say(f"You said: {text}")
                # engine.runAndWait()
                print(f"You said: {text}")
                return text

            except sr.RequestError as e:
                print("Could not request results:", e)

            except sr.UnknownValueError:
                engine.say("Sorry I couldn't get what you said")
                engine.runAndWait()
                print("Sorry I couldn't get what you said")

            except KeyboardInterrupt:
                engine.say("The program is stopping as per your command. Have a good day ahead")
                engine.runAndWait()
                print("The program is stopping as per your command. Have a good day ahead")
                exit(0)

def perform_task(command):
    command = command.lower()

    if "how are you" in command or "how have you been" in command:
        greetings = ["I'm fine. Thank you. What about you?", "I'm great. What about you?", "I'm good. How have you been?", "I'm fine. How about you?", "I'm great. How have you been?", "I'm great. How about you?"]
        resp = random.choice(greetings)
        engine.say(resp)
        engine.runAndWait()
        print(resp)

    elif "time" in command:
        curr_time = datetime.datetime.now().strftime("%H:%M")
        engine.say(f"The current time is {curr_time}")
        engine.runAndWait()
        print(f"The current time is {curr_time}")

    elif "date" in command:
        curr_date = datetime.datetime.now().strftime("%d-%B-%Y")
        engine.say(f"The current date is {curr_date}")
        engine.runAndWait()
        print(f"The current date is {curr_date}")

    elif "notepad" in command:
        subprocess.Popen("notepad.exe")
        engine.say("Opening notepad...")
        engine.runAndWait()
        print("Opening notepad...")

    elif "calculator" in command:
        subprocess.Popen("calc.exe")
        engine.say("Opening calculator...")
        engine.runAndWait()
        print("Opening calculator...")

    elif "chrome" in command:
        subprocess.Popen("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
        engine.say("Opening chrome browser...")
        engine.runAndWait()
        print("Opening browser...")
    
    elif "downloads" in command:
        subprocess.Popen(f'explorer "{os.path.join(os.path.expanduser("~"), "Downloads")}"')
        engine.say("Opening downloads...")
        engine.runAndWait()
        print("Opening downloads...")
    
    elif "one note" in command or "onenote" in command:
        subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\ONENOTE.EXE")
        engine.say("Opening OneNote...")
        engine.runAndWait()
        print("Opening OneNote...")

    elif "sublime" in command:
        subprocess.Popen("C:\\Program Files\\Sublime Text\\sublime_text.exe")
        engine.say("Opening Sublime Text...")
        engine.runAndWait()
        print("Opening Sublime Text...")
    
    elif "codeblocks" in command or "code blocks" in command:
        subprocess.Popen("C:\\Program Files\\CodeBlocks\\codeblocks.exe")
        engine.say("Opening codeblocks...")
        engine.runAndWait()
        print("Opening codeblocks...")
    
    elif "arduino" in command or "arduino ide" in command:
        subprocess.Popen("C:\\Program Files\\Arduino IDE\\Arduino IDE.exe")
        engine.say("Opening Arduino IDE...")
        engine.runAndWait()
        print("Opening Arduino IDE...")
    
    elif "excel" in command:
        subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE")
        engine.say("Opening Excel...")
        engine.runAndWait()
        print("Opening Excel...")
    
    elif "ms word" in command:
        subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE")
        engine.say("Opening MS Word...")
        engine.runAndWait()
        print("Opening MS Word...")
    
    elif "powerpoint" in command:
        subprocess.Popen("C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE")
        engine.say("Opening PowerPoint...")
        engine.runAndWait()
        print("Opening PowerPoint...")
    
    elif "ms edge" in command or "microsoft edge" in command:
        subprocess.Popen("C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe")
        engine.say("Opening Microsoft Edge...")
        engine.runAndWait()
        print("Opening Microsoft Edge...")
    
    elif "spotify" in command or "music" in command:
        subprocess.Popen("spotify.exe")
        engine.say("Opening spotify...")
        engine.runAndWait()
        print("Opening spotify...")
    
    elif "i want you to look up something for me" in command or "can you look for something" in command or "can you search something for me" in command:
        engine.say("Unfortunately, currently I'm not capable of doing that...")
        engine.runAndWait()

greet()

while True:
    command = listen()
    if command:
        perform_task(command)
    time.sleep(2)
