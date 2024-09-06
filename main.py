import time

import pyautogui
import speech_recognition as sr
import os

import wikipedia
import win32com.client
import random
speaker = win32com.client.Dispatch("SAPI.SpVoice")
import webbrowser
import openai
import subprocess
import datetime
import openai
import uuid
from requests import get
import pywhatkit

#openAI Api key
api_key = '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'

# -----------------------------for OpenAi-------------------------------------------------------
def ai(prompt):
    text=f"Open Ai response for Prompt : {prompt} \n******************************************\n\n"
    openai.api_key = api_key
    conversation_id = str(uuid.uuid4())
    # Your prompt for text completion
    prompt_text = prompt

    # Make a completion request
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt_text,
        max_tokens=100  # Adjust the number of tokens for completion
    )

    # Print the generated text
    print(response.choices[0].text.strip())

    text += response.choices[0].text.strip()
    speaker.Speak(text)
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"prompt-{random.randint(1,100)}", "w") as f:
        f.write(text)
# --------------------------------------for chatting -----------------------------------------------------------------
chatStr=""
def chat(query):
    global chatStr
    # print(chatStr)
    openai.api_key = api_key
    conversation_id = str(uuid.uuid4())
    # Your prompt for text completion
    chatStr +=query


    # Make a completion request
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=chatStr,
        max_tokens=100  # Adjust the number of tokens for completion
    )
    speaker.Speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

# --------------------------------for taking voice command (Mike input)-----------------------------------------------------------
def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening......")
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            print("Recognizing.......")
            query = r.recognize_google(audio, language="en-in")
            print(f"User Said : {query}")
            return query
        except Exception as e:
            return "Some Error Occurred . sorry from Jarvis  "

# ------------------------------------------wish function--------------------------------------------------------
def wish():
    hour= int(datetime.datetime.now().hour)

    if hour>=0 and hour<=12:
        speaker.Speak("Good Morning")
    elif hour>12 and hour<18:
        speaker.Speak("Good Afternoon")
    else:
        speaker.Speak("Good Evening")
    speaker.Speak("I am Jarvis sir. please tell me hoe can i help you")
# --------------------------------MAIN---------------------------------------------------------------------------
if __name__ == '__main__':
    wish()

    while True:
        # -------------------------------------------------------------------------
        query = takeCommand()
        # --------------------------------Notepad-----------------------------------------
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"]]

        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} sir....")
                webbrowser.open(site[1])

        if "the time" in query:
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            speaker.Speak(f"sir the time is {hour} o clock {min} minutes")

        elif "open notepad".lower() in query.lower():
            path="C:\\Windows\\system32\\notepad.exe"
            os.startfile(path)

        elif "close notepad".lower() in query.lower():
            speaker.Speak("Ok sir , Closing Notepad")
            os.system("taskkill /f /im  notepad.exe")

        elif "open command prompt".lower() in query.lower():
            os.system("start cmd")

        elif "close command prompt".lower() in query.lower():
            speaker.Speak("Ok sir , Closing ")
            os.system("taskkill /f /im  cmd.exe")

        elif "play some music" in query:

            speaker.Speak("Opening music sir....")
            music_dir = ("D:\\Songs")
            songs = os.listdir(music_dir)
            rd=random.choice(songs)
            os.startfile(os.path.join(music_dir,rd))

        elif "close music".lower() in query.lower():
            speaker.Speak("Ok sir , Closing music")
            os.system("taskkill /f /im  wmplayer.exe")

        elif "ip address".lower() in query.lower():
            ip=get('https://api.ipify.org').text
            print(f"IP : {ip}")
            speaker.Speak(f"your IP address is {ip}")

        elif "wikipedia".lower() in query.lower():
            speaker.Speak("Searching Wikipedia...")
            query=query.replace("wikipedia","")
            results = wikipedia.summary(query,sentences=2)
            speaker.Speak("According to wikipedia ")
            print(results)
            speaker.Speak(results)

        elif "search on google".lower() in query.lower():
            speaker.Speak("sir, What should i search on google")
            cm=takeCommand().lower()
            webbrowser.open(f"{cm}")

        elif "search on youtube".lower() in query.lower():
            speaker.Speak("sir, What should i search on youtube")
            cm = takeCommand().lower()
            pywhatkit.playonyt(f"{cm}")

        elif "open my Instagram Account".lower() in query.lower():
            speaker.Speak("Opening Sir......")
            webbrowser.open("https://www.instagram.com/aniket_ahire__0401/")

        elif "Using Artificial Intelligence".lower() in query.lower():
            ai(prompt=query)

        elif "switch the window".lower() in query.lower():
            pyautogui.keyDown("alt")
            pyautogui.press("tab")
            time.sleep(1)
            pyautogui.keyUp("alt")

        elif "Thank You jarvis".lower() in query.lower():
            speaker.Speak("My Pleasure sir ")
            exit()

        else:
            chat(query)
