import speech_recognition as sr 
import pyttsx3 
import time
import pywhatkit
import pyautogui
import pyjokes
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def sound(n):
    speech = pyttsx3.init()     
    speech.setProperty('rate', 180)     
    speech.say(n)           
    speech.runAndWait()

def command():
    r = sr.Recognizer() 
    with sr.Microphone() as source:
        audio=r.listen(source)
    try:
        audio=r.recognize_google(audio).lower()
        print(audio)
    except:
        audio=''
    return audio

def greeting():
    hour=int(time.strftime("%H"))
    if hour>3 and hour<12:
        welcome_text='Good Morning'
    elif hour>=12 and hour<17:
        welcome_text='Good Afternoon'
    else:
        welcome_text='Good Evening'
    welcome_text+=', BOT-TOP here.'
    print(welcome_text)
    sound(welcome_text)


audio=''
while 'stop the program' not in audio: 
    audio=command()
    if 'wake up' in audio:
        greeting()
        while 'sleep' not in audio:
            audio=command() 
            if 'play' in audio and 'youtube' in audio:
                audio=audio.replace('play','').replace('youtube','')
                print('Starting youtube')
                sound('Starting youtube')
                pywhatkit.playonyt(audio)

            elif 'open' in audio:
                audio=audio.replace('open','')
                sound('opening...')
                print('opening...')
                pyautogui.press('super')
                pyautogui.typewrite(audio)
                pyautogui.press('enter')

            elif 'joke' in audio:
                joke=pyjokes.get_joke()
                joke+='  Haa-haa-haa'
                print(joke)
                sound(joke) 

            elif 'type' in audio:
                audio=audio.replace('type','')
                print(audio)
                sound('Typing...')
                pyautogui.typewrite(audio)

            elif 'minimise' in audio:
                pyautogui.hotkey('win','down','down')

            elif 'maximize' in audio:
                pyautogui.hotkey('win','up','up')

            elif 'close all' in audio:
                pyautogui.hotkey('alt','f4')
            
            elif 'close' in audio:
                pyautogui.hotkey('ctrl','w')

            elif 'search' in audio:
                audio=audio.replace('search','')
                pywhatkit.search(audio)

            elif 'press enter' in audio:
                pyautogui.press('enter')

            elif 'space bar' in audio:
                pyautogui.press('space')

            elif 'go left' in audio:
                audio=audio.split()
                for i in audio:
                    if i.isdigit():
                        pyautogui.press('left',presses=int(i))
                        break
                else:
                    pyautogui.press('left')

            elif 'go right' in audio:
                audio=audio.split()
                for i in audio:
                    if i.isdigit():
                        pyautogui.press('right',presses=int(i))
                        break
                else:
                    pyautogui.press('right')

            elif 'go write' in audio:
                audio=audio.split()
                for i in audio:
                    if i.isdigit():
                        pyautogui.press('right',presses=int(i))
                        break
                else:
                    pyautogui.press('right')                
            
            elif 'go up' in audio:
                audio=audio.split()
                for i in audio:
                    if i.isdigit():
                        pyautogui.press('up',presses=int(i))
                        break
                else:
                    pyautogui.press('up')

            elif 'go' and 'down' in audio:
                audio=audio.split()
                for i in audio:
                    if i.isdigit():
                        pyautogui.press('down',presses=int(i))
                        break
                else:
                    pyautogui.press('down')

            elif 'godown' in audio:
                audio=audio.split()
                for i in audio:
                    if i.isdigit():
                        pyautogui.press('down',presses=int(i))
                        break
                else:
                    pyautogui.press('down')

            elif 'save file' in audio:
                pyautogui.hotkey('ctrl','s')

            elif 'select all' in audio:
                pyautogui.hotkey('ctrl','a')

            elif 'undo' in audio:
                pyautogui.hotkey('ctrl','z')

            elif 'redo' in audio:
                pyautogui.hotkey('ctrl','y')

            elif 'delete' in audio:
                pyautogui.hotkey('del')

            elif 'set volume' in audio:
                audio=audio.split()
                for i in audio:
                    if '%' in i:
                        print(i)
                        i=i.replace('%','')
                        device=AudioUtilities.GetSpeakers()
                        interface=device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        volume=interface.QueryInterface(IAudioEndpointVolume)
                        volume.SetMasterVolumeLevelScalar(int(i)/100,None)
                        break

            elif 'reload' in audio:
                pyautogui.hotkey('ctrl','r')

            elif 'play' in audio:
                pyautogui.hotkey('playpause')

            elif 'pause' in audio:
                pyautogui.hotkey('playpause')

            elif 'full screen' in audio:
                pyautogui.hotkey('f11')

            elif 'stop the program' in audio:
                break

        else:
            print('\a')

else:
    print('Thank you so much, have a good day.')
    sound('Thank you so much, have a good day.')