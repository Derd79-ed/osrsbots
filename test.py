import win32gui
import cv2
import time
import win32com.client
import pyautogui
from pyautogui import keyDown, keyUp
import random
import sys

def find_window():
    #Look through all your windows
    def search(handle, window):
        text = win32gui.GetWindowText(handle)
        rect = win32gui.GetWindowRect(handle)
        #Get coordinates
        x = rect[0]
        y = rect[1]
        w = rect[2] - x
        h = rect[3] - y
        #If the window has Runelite, it saves the coordinates
        if "runelite" in text.lower():
            print(f"Window {win32gui.GetWindowText(handle)}:")
            print(f"\tLocation: ({x},{y})")
            print(f"\tSize: ({w},{h})")
            window.append({'handle':handle,'x':x,'y':y,'w':w,'h':h})
    #Store all windows that match and output found window, will be first in list (and only)
    window = []
    win32gui.EnumWindows(search, window)
    return window[0]

#Get window dimensions
window = find_window()
print(window)
x,y,w,h = window['x'], window['y'], window['w'], window['h']
print(x,y,w,h)

#Set window into foreground
shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys('%')

win32gui.SetForegroundWindow(window['handle'])
time.sleep(1)

#Initialized PyAutoGui    
pyautogui.FAILSAFE = True
  
# ---------- VARIABLES SECTION ----------

#Display tab coordinates
display_tab = x+574, y+256

#Brightness level coordinates
brightness = x+711, y+333

# Coordinates for Hitpoint healer
hp_x = x + 550
hp_y = y + 85

offset_x = random.randint(-5, 5)
offset_y = random.randint(-5, 5)

#Countdown timer
print("Starting", end="")
for i in range(0,3):
    print(".", end="")
    time.sleep(1)
print("Go")

#needed for consistent clicks
def long_click(input_x,input_y=None,button="left"):
    if input_y is None:
        input_x, input_y = input_x #Unpack tuple result of pyautogui.locate
    pyautogui.moveTo(input_x,input_y)
    pyautogui.mouseDown(button=button)
    time.sleep(0.4)
    pyautogui.mouseUp(button=button)

# ---------- SETTING THE BASICS START ----------

#Click on compass
long_click(x+564, y+48)

#Move Mouse in center of window
pyautogui.moveTo(x+w/2, y+h/2)

#Move camera up
keyDown("up")
time.sleep(1.5)
keyUp("up")

# Click on setting tab
#keyDown("F10")
#time.sleep(0.5)

#Set the brightness level
#long_click(brightness, None)
#time.sleep(0.5)

# click on inventory tab
keyDown("F1")
time.sleep(0.5)

#Move Mouse in center of window
pyautogui.moveTo(x+w/2, y+h/2)
time.sleep(0.2)

#Scroll up a lot
for i in range(10):
    pyautogui.scroll(-550)
    time.sleep(0.1)

#Scroll up a lot
for i in range(10):
    pyautogui.scroll(250)
    time.sleep(0.1)

# ---------- SETTING THE BASICS END ----------

def bank_wood():
    pyautogui.moveTo(1545, 227, duration=1.3)
    long_click(1545,227)
    time.sleep(7)
    pyautogui.moveTo(1112,353,duration=0.9)
    long_click(1112,353)
    time.sleep(2)
    bank = pyautogui.locateCenterOnScreen('deposit.png', confidence=0.8)
    if bank:
        pyautogui.moveTo(bank,None,duration=0.9)
        long_click(bank)

    pyautogui.moveTo(1473,312,duration=1.8)
    long_click(1473,312)
    time.sleep(7.3)


def cut_wood():
    try:
        # Check if we are already cutting (by finding the "cutting" image)
        isCutting = bool(pyautogui.locateOnScreen('cutting.png', confidence=0.5))
    except pyautogui.ImageNotFoundException:
        isCutting = False

    if isCutting:
        print("Already cutting. Skipping action.")
        return  # Exit early if already cutting

    # Otherwise, try to find a willow tree and cut it
    try:
        willow1 = pyautogui.locateCenterOnScreen('willow1.png', confidence=0.8)
    except pyautogui.ImageNotFoundException:
        willow1 = None

    if willow1:
        long_click(willow1)
        return

    try:
        willow2 = pyautogui.locateCenterOnScreen('willow2.png', confidence=0.8)
    except pyautogui.ImageNotFoundException:
        willow2 = None

    if willow2:
        long_click(willow2)
    else:
        print("Neither willow1 nor willow2 was found. Skipping.")

def is_inventory_full():
    try:
        return bool(pyautogui.locateOnScreen('inventory_full.png',confidence=0.99))
    except pyautogui.ImageNotFoundException:
        return False
# === Main Loop ===
while True:
    if is_inventory_full():
        bank_wood()
    else:
        cut_wood()
    
    time.sleep(3)