import pyautogui as p
import os
import time
from pywinauto import application
from selenium import webdriver
import speech_recognition as s
import win32com.client as wincl




def install_py_mod(s):
    try:
        app=application.Application()
        app.start("cmd.exe")
    except:
        time.sleep(3)
        p.typewrite("pip install "+s)
        p.press("enter")
        p.typewrite("exit()")
        time.sleep(2)
        p.press("enter")

def facebook():
    driver=webdriver.Chrome('C:\\Users\\user\\Desktop\\auto\\chromedriver.exe')
    usr="xxx"
    pwd="xxx"
    driver.get('https://www.facebook.com/')
    for i in range(3):
        driver.refresh()
    username_box=driver.find_element_by_id('email')
    username_box.send_keys(usr)
    pass_box=driver.find_element_by_id('pass')
    pass_box.send_keys(pwd)
    login_box=driver.find_element_by_css_selector("#u_0_3")
    login_box.submit()

def chrome():
    p.press('win')
    p.typewrite("chrome")
    time.sleep(2)
    p.press("enter")

def youtube():
    p.press('win')
    p.typewrite("chrome")
    time.sleep(2)
    p.press("enter")
    time.sleep(2)
    p.typewrite("y")
    time.sleep(1)
    p.press("enter")
    
    

def dc():
    p.press('win')
    p.typewrite("dc")
    time.sleep(2)
    p.press("enter")
    time.sleep(4)
    p.typewrite("xxx")
    p.press("enter")


    
    
def notepad():
    p.press('win')
    p.typewrite("notepad")
    time.sleep(2)
    p.press("enter")

i=1
def screenshot():
    p.screenshot('my_pic'+str(i)+'.png')

def hotspot():
    p.moveTo(1637,1052)
    p.click()
    time.sleep(2)
    p.moveTo(1737,977)
    p.click()
    
    
def placement():
    usr='xxx'
    pwd='xxx'
    driver=webdriver.Chrome('C:\\Users\\user\\Desktop\\auto\\chromedriver.exe')
    driver.get('http://placement.bitmesra.ac.in/')
    for i in range(3):
        driver.refresh()
    username_box=driver.find_element_by_id('txtUsername')
    username_box.send_keys(usr)
    pass_box=driver.find_element_by_id('txtPassword')
    pass_box.send_keys(pwd)
    login_box=driver.find_element_by_id('imgSubmit')
    login_box.click()
    driver.get('http://placement.bitmesra.ac.in/Student/Jobs.aspx')

def codechef():
    usr="xxx"
    pwd='xxx'
    driver=webdriver.Chrome('C:\\Users\\xxx\\Desktop\\auto\\chromedriver.exe')
    driver.get('https://www.codechef.com/')
    driver.refresh();
    for i in range(2):
        driver.refresh()
    username_box=driver.find_element_by_id("edit-name")
    username_box.send_keys(usr)
    pass_box=driver.find_element_by_id("edit-pass")
    pass_box.send_keys(pwd)
    login_box=driver.find_element_by_id('edit-submit')
    login_box.click()
    driver.get('https://www.codechef.com/contests')

#https://github.com/gokulraviteja
def github():
    driver=webdriver.Chrome('C:\\Users\\xxx\\Desktop\\auto\\chromedriver.exe')
    driver.get('https://github.com/xxx')
    for i in range(2):
        driver.refresh()
    

def downloads():
    p.press("win")
    p.typewrite("file")
    time.sleep(2)
    p.press('enter')
    time.sleep(2)
    p.press("down")
    p.press("enter")
    
def cyber():
    p.press("win")
    p.typewrite("edge")
    time.sleep(2)
    p.press("enter")
    time.sleep(2)
    p.typewrite("https://172.16.1.1:8090/")
    time.sleep(2)
    p.press("enter")
    time.sleep(2)
    p.moveTo(780,444)
    p.click()
    time.sleep(2)
    p.moveTo(817,678)
    p.click()
    time.sleep(2)
    p.moveTo(761,393)
    p.click()
    time.sleep(2)
    p.moveTo(759,441)
    p.click()
    time.sleep(2)
    p.moveTo(749,498)
    p.click()
    p.moveTo(1882,15)
    p.click()
    time.sleep(2)
    
    
    
    

while(1):

    r=s.Recognizer()
    with s.Microphone() as source:
        print("START")
        speak = wincl.Dispatch("SAPI.SpVoice")
        speak.Speak("Start")
        audio=r.listen(source)
        print("---")
        #print("**")
    try:
        ll=r.recognize_google(audio)
        print("ds")
        print("YOU SAID : "+ll);
        speak = wincl.Dispatch("SAPI.SpVoice")
        
        if('hello' in ll or 'Hello' in ll):
            speak.Speak("HELLO GOKUL!!")
           
        elif('DC' in ll):
            speak.Speak("OPENING DC MATE")
            print("OPENING DC MATE");
            dc()
        elif('Facebook' in ll):
            speak.Speak("OPENING FACEBOOK MATE")
            print("OPENING FACEBOOK MATE");
            facebook()
        elif('codechef' in ll):
            speak.Speak("OPENING CODECHEF MATE")
            print("OPENING CODECHEF MATE");
            codechef()
        elif('placement' in ll):
            speak.Speak("OPENING PLACEMENT MATE")
            print("OPENING PLACEMENT SITE MATE");
            placement()
        elif('Python' in ll or 'install' in ll):
            ll=ll.split()
            speak.Speak(" Installing "+ll[-1]+" MATE")
            print("INSTALLING " + ll[-1])
            install_py_mod(ll[-1])
        elif('Notepad' in ll):
            speak.Speak(" OPENING NOTEPAD MATE")
            print("OPENING NOTEPAD")
            notepad()
        elif("GitHub" in ll):
            speak.speak("OPENING GITHUB MATE")
            print("OPENING GITHUB")
            github()
        elif("Chrome" in ll):
            speak.speak("OPENING CHROME MATE")
            chrome()
        elif("download" in ll):
            speak.speak("OPENING DOWNLOADS MATE")
            downloads()
        elif("screen" in ll or "shot" in ll):
            speak.speak("TAKING SCREENSHOT MAKE IT READY MATE")
            time.sleep(2)
            screenshot()
        elif("hotspot" in ll):
            speak.speak("Turning on Hotspot")
            hotspot()
        elif("cyberoam" in ll):
            speak.speak("LOGGING IN TO CYBERROAM MATE")
            cyber()
        elif("YouTube" in ll or "Tube" in ll or "You" in ll):
            speak.speak("OPENING YOUTUBE MATE")
            youtube()
            
            

            
    except:
        pass;























"""
import ctypes, sys
#p.FAILSAFE=True
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
if is_admin():
    app=application.Application()
    app.start("cmd.exe")
else:
    # Re-run the program with admin rights
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    app=application.Application()
    app.start("cmd.exe")
"""












