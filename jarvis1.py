from tkinter import *
import cv2
import PIL.Image
import PIL.ImageTk
import subprocess
import wolframalpha
import pyttsx3
import operator
import speech_recognition as sr
from pynput.keyboard import Key, Controller
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import ctypes
import time
import shutil
from PIL import ImageGrab
import pyautogui
import roman
from PIL import Image


app_id = 'Wolframalpha-key'

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


window = Tk()

window.resizable(FALSE, FALSE)

global var
global var1
global uname
global assname2

assname2 = ("Jarvis 1 point o")

var = StringVar()
var1 = StringVar()

keyboard = Controller()


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        var.set("Good Morning Sir !")
        window.update()
        speak("Good Morning Sir !")

    elif hour >= 12 and hour < 18:
        var.set("Good Afternoon Sir !")
        window.update()
        speak("Good Afternoon Sir !")

    else:
        var.set("Good Evening Sir !")
        window.update()
        speak("Good Evening Sir !")

    assname2 = ("Jarvis 1 point o")
    speak("I am your Desktop Voice Assistant")
    speak("My nick name is ")
    speak(assname2)


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        var.set("Listening...")
        window.update()
        print("Listening...")
        r.pause_threshold = 1
        r.energy_threshold = 400
        audio = r.listen(source)

    try:
        var.set("Recognizing...")
        window.update()
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to Recognizing your voice.")
        return "None"
    var.set(query)
    window.update()
    return query


def uName():
    var.set("What should i call you sir")
    window.update()
    speak("What should i call you sir")
    uname = takeCommand()
    speak("Welcome Mister")
    speak(uname)
    speak("How can i Help you, Sir")


def play():
    btn0['state'] = 'disabled'
    wishMe()
    uName()

    while True:
        query = takeCommand().lower()
        k = 1
        s1 = 1
 # basic questions
        if 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Sir")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that your fine")

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak("Jarvis 1 point o")

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Ranjeet Sir.")

        elif "what you do it" in query or "what are you doing" in query or "which type of command" in query or "what to do" in query or "which type of taks" in query:
            speak("Sir I am doing various task such as open google, facebook ,youtube,instagram, whatsapp,gmail,google chrome , take a photo, search wikipedia to abstaract the required data,and can answer computational and geographical questions too.")

        elif "who i am" in query:
            speak("If you talk then definately your human.")

        elif "why you came to world" in query:
            speak("Thanks to Ranjeet. further It's a secret")

        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")

        elif "who are you" in query:
            speak("I am your desktop voice assistant created by Ranjeet Sir")

        elif 'reason for you' in query:
            speak("I was created as a Major project by Mister Ranjeet")

        elif "Good Morning" in query:
            speak("A warm" + query)
            speak("How are you Mister")

        elif 'date' in query:
            date = datetime.datetime.now().strftime("%d: %m: %Y")
            print(date)
            speak(date)

        elif "hello" in query:
            print("hello sir")
            speak("Hello sir !")

        elif "thank you" in query:
            speak('You are welcome sir !')

  # most asked question from google Assistant
        elif "will you be my gf" in query or "will you be my bf" in query:
            speak("I'm not sure about, may be you should give me some time")

        elif "how are you" in query:
            speak("I'm fine, glad you me that")

        elif "i love you" in query:
            speak("It's hard to understand")

        elif "what is" in query:
            client = wolframalpha.Client("Wolframalpha-key")
            res = client.query(query)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")

  # Open and Close Website
        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open stackoverflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")

        elif 'close youtube' in query or 'close google' in query or 'close stackoverflow' in query or 'close wikipedia' in query:
            os.system("taskkill /im iexplore.exe /f")
            speak("Ok Sir")

        elif 'open wikipedia' in query:
            webbrowser.open("wikipedia.com")

        elif 'open chrome' in query or 'open browser' in query:
            codePath = "C:\\Program Files\\Google\\Chrome\\Application\\Chrome.exe"
            os.startfile(codePath)
            speak("Opening the browser Sir")

        elif 'open facebook' in query:
            speak("Here you go to facebook")
            webbrowser.open("facebook.com")

        elif 'open instagram' in query:
            speak("Here you go to Instagram")
            webbrowser.open("instagram.com")

        elif 'open whatsapp' in query:
            speak("here you go to whatsapp web")
            webbrowser.open("web.whatsapp.com")

        elif 'close facebook' in query or 'close instagram' in query or 'close whatsapp' in query:
            os.system("taskkill /im iexplore.exe /f")
            speak("Ok Sir")

 # Open and Close Microsoft Office Word , Powerpoint and Excel

        elif 'open powerpoint' in query:
            speak("opening Power Point presentation")
            power = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Microsoft Office PowerPoint 2007.lnk"
            os.startfile(power)

        elif 'open word' in query:
            speak("opening Microsoft Office Word")
            power2 = "C:\\Program Files\\Microsoft Office\\Office12\\WINWORD.exe"
            os.startfile(power2)
            w = 1
            while(w == 1):
                speak("Would you like to write sir")
                n2 = takeCommand()
                if 'yes' in n2:
                    speak("sir please tell me , what should i write in ms word sir")
                    pr2 = takeCommand()
                    pyautogui.typewrite(pr2)
                else:
                    w = 2
                    speak("ok sir")

        elif 'open excel' in query:
            speak("opening Power Point presentation")
            power3 = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Microsoft Office Excel 2007.lnk"
            os.startfile(power3)

        elif 'close word' in query:
            os.system("taskkill /im  WINWORD.exe /f")

        elif 'close powerpoint' in query:
            os.system("taskkill /im  POWERPNT.exe /f")

        elif 'close excel' in query:
            os.system("taskkill /im  EXCEL.exe /f")

 # search in Wikipedia
        elif 'in wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

 # Open and Write Note and text in notepad
        elif 'open notepad' in query:
            subprocess.Popen('notepad.exe')
            k = 2
            while(k == 2):
                speak("Would you like to write sir")
                n = takeCommand()
                if 'yes' in n:
                    speak("sir please tell me , what i write in notpad")
                    pr = takeCommand()
                    pyautogui.typewrite(pr)
                else:
                    k = 3
                    speak("ok sir")

        elif "write a note" in query or 'write a diary' in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)

        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(30))
            s1 = 2

        elif 'close notepad' in query or 'close note' in query:
            if (k == 2) or (k == 3) or (s1 == 2):
                os.system("taskkill /im notepad.exe /f")
            else:
                speak(
                    "Sorry sir , not any notepad application is open , so i can not be closed it")

  # Changing the command

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            usrName = query

        elif "change your name" in query:
            speak("What would you like to call me, Sir ")
            assname2 = takeCommand()
            speak("Thanks for naming me")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(
                20, 0, "Location of wallpaper", 0)
            speak("Background changed succesfully")

 # System Command and various other command

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()

        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif 'clean recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Cleaned , Sir")

        elif "restart" in query:
            speak("ok Sir")
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "don't listen" in query or "stop listening" in query:
            speak("stop jarvis from listening commands")
            btn0['state'] = 'normal'
            x = 2
            break

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            speak("I am open the map")
            webbrowser.open(
                "https://www.google.com/maps/place/" + location + "")

        elif 'close map' in query:
            speak("ok Sir")
            os.system("taskkill /im iexplore.exe /f")

        elif 'exit' in query or 'quit' in query or 'bye' in query:
            speak("Thanks for giving me your time")
            speak("ok bye , Sir")
            exit()
            window.destroy()

        elif 'joke' in query:
            speak("ok Sir")
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'who is' in query:
            results = wikipedia.summary(query, sentences=2)
            speak(results)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir , the time is {strTime}")

        elif 'find in laptop' in query:
            speak("what is find and open , Sir")
            f = takeCommand().lower()
            pyautogui.press('win', interval=0.2)
            pyautogui.typewrite(f, interval=0.4)
            pyautogui.press('enter', interval=0.4)

        elif 'take screenshot' in query:
            image = ImageGrab.grab(bbox=(0, 0, 1366, 768))
            image.save("./img.png")
            speak("ok Sir")

        elif 'show screenshot' in query:
            speak("ok Sir")
            os.startfile("./img.png")

        elif 'close picture' in query or 'close photo' in query or 'close screenshot' in query:
            speak("ok Sir")
            if subprocess.Popen('photos.exe') == True:
                os.system("taskkill /im photos.exe /f")
            else:
                speak("Sorry sir ,none picture is open , so i can not close it ")

 # close anything
        elif 'close' in query:
            speak("Which application  I closed it sir")
            m1 = takeCommand().lower()
            a = m1.join('.exe')
            if not a:
                speak("Sorry sir ,not any" + m1 +
                      " is open , so i can not close it ")
            else:
                speak("Ok Sir , I am closing " + m1)
                os.system("taskkill /im "+m1+".exe /f")

 # Open Drive
        elif 'open drive' in query:
            speak("which drive Sir")
            dr = takeCommand()
            if 'c' in dr:
                speak("Ok sir I am opening C drive")
                os.startfile("C:\\")
            elif 'f' in dr:
                speak("Ok sir I am opening F drive")
                os.startfile("F:\\")
            elif 'e' in dr:
                speak("Ok sir I am opening E drive")
            else:
                speak("This drive not in your computer ")

        elif 'click photo' in query:
            speak("ok Sir")
            stream = cv2.VideoCapture(0)
            grabbed, frame = stream.read()
            if grabbed:
                cv2.imshow('pic', frame)
                cv2.imwrite('pic.jpg', frame)
            stream.release()

        elif 'record video' in query:
            speak("ok Sir")
            cap = cv2.VideoCapture(0)
            out = cv2.VideoWriter('output.avi', -1, 20.0, (640, 480))
            while(cap.isOpened()):
                ret, frame = cap.read()
                if ret:

                    out.write(frame)

                    cv2.imshow('frame', frame)
                    if cv2.waitKey(1) & 0xFF == ord('q'):
                        break
                else:
                    break
            cap.release()
            out.release()
            cv2.destroyAllWindows()

        if 'type' in query:
            speak("tell me a sentence")
            sentences = takeCommand().lower()
            keyboard.type(sentences)
            speak("ok Sir")

        elif 'select all' in query:
            keyboard.press(Key.ctrl_l)
            keyboard.press('a')
            speak("ok Sir")

        elif 'copy' in query:
            keyboard.press(key.ctrl_l)
            keyboard.press('c')
            speak("ok Sir")

        elif 'cut' in query:
            keyboard.press(key.ctrl_l)
            keyboard.press('x')
            speak("ok Sir")

        elif 'paste' in query:
            keyboard.press(key.ctrl_l)
            keyboard.press('v')
            speak("ok Sir")

        elif 'undo' in query:
            keyboard.press(Key.ctrl_l)
            keyboard.press('z')
            speak("ok Sir")

        elif 'enter' in query:
            keyboard.press(Key.enter)
            speak("ok Sir")

        elif 'backspace' in query:
            keyboard.press(Key.backspace)
            speak("ok Sir")

        elif 'select back' in query:
            keyboard.press(Key.shift_l)
            keyboard.press(Key.left)
            speak("ok Sir")

        elif 'selct forward' in query:
            keyboard.press(Key.shift_l)
            keyboard.press(Key.right)
            speak("ok Sir")

        elif 'backward' in query:
            keyboard.press(Key.left)
            speak("ok Sir")

        elif 'forward' in query:
            keyboard.press(Key.right)
            speak("ok Sir")

        elif 'top' in query:
            keyboard.press(Key.up)
            speak("ok Sir")

        elif 'bottom' in query:
            keyboard.press(Key.down)
            speak("ok Sir")

        elif 'delete' in query:
            keyboard.press(Key.delete)
            speak("ok Sir")

        elif query == 'none':
            speak("what is your next command sir")
            continue

        elif 'search in google' in query:
            speak("what i search it Sir")
            query = takeCommand().lower()
            temp = query.replace('', '+')
            g_url = "https://www.google.com/search?q="
            webbrowser.open(g_url, temp)

        elif 'search in youtube' in query:
            speak("what i search it Sir")
            query = takeCommand().lower()
            temp = query.replace('', '+')
            g_url = "https://www.youtube.com/results?search_query="
            webbrowser.open(g_url, temp)

        speak("what is your next command sir")


def update(ind):
    frame = frames[(ind) % 100]
    ind += 1
    label.configure(image=frame)
    window.after(100, update, ind)


label2 = Label(window, textvariable=var1, bg='#FAB60C')
label2.config(font=("Courier", 20))
var1.set('User Said:')
label2.pack()

label1 = Label(window, textvariable=var, bg='#ADD8E6')
label1.config(font=("Courier", 20))
var.set('Welcome')
label1.pack()

frames = [PhotoImage(file='Assistant.gif', format='gif -index %i' % (i))
          for i in range(100)]
window.title('Desktop Voice Assistant')

label = Label(window, width=500, height=500)
label.pack()
window.after(0, update, 0)

btn0 = Button(text='START', width=20, command=play, bg='#5C85FB')
btn0.config(font=("Courier", 12))
btn0.pack()

btn2 = Button(text='EXIT', width=20, command=window.destroy, bg='#5C85FB')
btn2.config(font=("Courier", 12))
btn2.pack()

speak("Welcome Sir")

window.mainloop()

speak("Thank you Sir")
