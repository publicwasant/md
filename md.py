from PIL import ImageGrab
from numpy import *

import win32api
import win32con
import win32gui
import atexit
import pygetwindow
import pytesseract

import time
import math

windows = None
inputTimeDelay = 0.4
catchRetryTimeDelay = 1.2
pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"
vkCode = {
    'backspace':0x08,
    'tab':0x09,
    'clear':0x0C,
    'enter':0x0D,
    'shift':0x10,
    'ctrl':0x11,
    'alt':0x12,
    'pause':0x13,
    'caps_lock':0x14,
    'esc':0x1B,
    'spacebar':0x20,
    'page_up':0x21,
    'page_down':0x22,
    'end':0x23,
    'home':0x24,
    'left_arrow':0x25,
    'up_arrow':0x26,
    'right_arrow':0x27,
    'down_arrow':0x28,
    'select':0x29,
    'print':0x2A,
    'execute':0x2B,
    'print_screen':0x2C,
    'ins':0x2D,
    'del':0x2E,
    'help':0x2F,
    '0':0x30,
    '1':0x31,
    '2':0x32,
    '3':0x33,
    '4':0x34,
    '5':0x35,
    '6':0x36,
    '7':0x37,
    '8':0x38,
    '9':0x39,
    'a':0x41,
    'b':0x42,
    'c':0x43,
    'd':0x44,
    'e':0x45,
    'f':0x46,
    'g':0x47,
    'h':0x48,
    'i':0x49,
    'j':0x4A,
    'k':0x4B,
    'l':0x4C,
    'm':0x4D,
    'n':0x4E,
    'o':0x4F,
    'p':0x50,
    'q':0x51,
    'r':0x52,
    's':0x53,
    't':0x54,
    'u':0x55,
    'v':0x56,
    'w':0x57,
    'x':0x58,
    'y':0x59,
    'z':0x5A,
    'numpad_0':0x60,
    'numpad_1':0x61,
    'numpad_2':0x62,
    'numpad_3':0x63,
    'numpad_4':0x64,
    'numpad_5':0x65,
    'numpad_6':0x66,
    'numpad_7':0x67,
    'numpad_8':0x68,
    'numpad_9':0x69,
    'multiply_key':0x6A,
    'add_key':0x6B,
    'separator_key':0x6C,
    'subtract_key':0x6D,
    'decimal_key':0x6E,
    'divide_key':0x6F,
    'F1':0x70,
    'F2':0x71,
    'F3':0x72,
    'F4':0x73,
    'F5':0x74,
    'F6':0x75,
    'F7':0x76,
    'F8':0x77,
    'F9':0x78,
    'F10':0x79,
    'F11':0x7A,
    'F12':0x7B,
    'F13':0x7C,
    'F14':0x7D,
    'F15':0x7E,
    'F16':0x7F,
    'F17':0x80,
    'F18':0x81,
    'F19':0x82,
    'F20':0x83,
    'F21':0x84,
    'F22':0x85,
    'F23':0x86,
    'F24':0x87,
    'num_lock':0x90,
    'scroll_lock':0x91,
    'left_shift':0xA0,
    'right_shift ':0xA1,
    'left_control':0xA2,
    'right_control':0xA3,
    'left_menu':0xA4,
    'right_menu':0xA5,
    'browser_back':0xA6,
    'browser_forward':0xA7,
    'browser_refresh':0xA8,
    'browser_stop':0xA9,
    'browser_search':0xAA,
    'browser_favorites':0xAB,
    'browser_start_and_home':0xAC,
    'volume_mute':0xAD,
    'volume_Down':0xAE,
    'volume_up':0xAF,
    'next_track':0xB0,
    'previous_track':0xB1,
    'stop_media':0xB2,
    'play/pause_media':0xB3,
    'start_mail':0xB4,
    'select_media':0xB5,
    'start_application_1':0xB6,
    'start_application_2':0xB7,
    'attn_key':0xF6,
    'crsel_key':0xF7,
    'exsel_key':0xF8,
    'play_key':0xFA,
    'zoom_key':0xFB,
    'clear_key':0xFE,
    '+':0xBB,
    ',':0xBC,
    '-':0xBD,
    '.':0xBE,
    '/':0xBF,
    '`':0xC0,
    ';':0xBA,
    '[':0xDB,
    '\\':0xDC,
    ']':0xDD,
    "'":0xDE,
    '`':0xC0}

def catch(title, timeDelay=0.5, retry=0):
    global windows

    while windows is None:
        if title in pygetwindow.getAllTitles():
            windows = pygetwindow.getWindowsWithTitle(title)[0]
        else:
            windows = None
        time.sleep(timeDelay)
    else:
        try:
            x, y = windows.topleft
            width, height = windows.size

            win32gui.SetWindowPos(windows._hWnd, win32con.HWND_TOPMOST, x, y, width, height, 0)
            win32gui.SetForegroundWindow(windows._hWnd)
        except:
            print("catch: retry " + str(retry))
            time.sleep(catchRetryTimeDelay)
            catch(title, timeDelay, retry + 1)

def release():
    if windows is not None:
        x, y = windows.topleft
        width, height = windows.size
        win32gui.SetWindowPos(windows._hWnd, win32con.HWND_NOTOPMOST, x, y, width, height, 0)

def pos(ps):
    xp, yp = ps
    x1, y1 = windows.topleft
    w, h = windows.size
        
    return (
        math.floor(x1 + (w*xp)), 
        math.floor(y1 + (h*yp)))

def cap(c1=(0, 0), c2=(0, 0)):
    x1, y1 = windows.topleft
    x2, y2 = windows.bottomright
    w, h = windows.size
    x1c, y1c = c1
    x2c, y2c = c2

    return ImageGrab.grab((
        math.floor(x1 + (w*x1c)), 
        math.floor(y1 + (h*y1c)), 
        math.floor(x2 - (w*x2c)), 
        math.floor(y2 - (h*y2c))))

def text(c1=(0, 0), c2=(0, 0)):
    str = pytesseract.image_to_string(cap(c1, c2))
    str = str.replace("\x0c", "")
    str = str.replace("\n", "/")
    str = str.replace(" ", "/")

    return [s for s in str.split("/") if s != ""]

def click(ps):
    win32api.SetCursorPos(pos(ps))
    time.sleep(inputTimeDelay)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(inputTimeDelay)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(inputTimeDelay)

def key(key):
    time.sleep(inputTimeDelay)
    win32api.keybd_event(vkCode[key], 0, 0, 0)
    time.sleep(inputTimeDelay)
    win32api.keybd_event(vkCode[key], 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(inputTimeDelay)

atexit.register(release)

class SimpleBot(object):
    def __init__(self):
        self.title = "Mir4G[0]"
        self.missionAcceptPos = (.92, .39)
        self.missionAutoPos = (.92, .39)
        self.missionClaimPos1 = (.78, .335)
        self.missionClaimPos2 = (.035, .59)

    def open(self):
        catch(self.title)
        print("open: catch=" + self.title)
        
        key("esc")
        print("open: key=esc")

        key("F6")
        print("open: key=f6")

    def close(self):
        catch(self.title)
        print("close: catch=" + self.title)
        
        key("esc")
        print("close: key=esc")

    def accept(self):
        catch(self.title)
        print("accept: catch=" + self.title)

        click(self.missionAcceptPos)
        print("accept: click=" + str(self.missionAcceptPos))
    
    def auto(self):
        catch(self.title)
        print("auto: catch=" + self.title)

        click(self.missionAutoPos)
        print("auto: click=" + str(self.missionAutoPos))

        count = 0

        while True:
            claimText = text(self.missionClaimPos1, self.missionClaimPos2)
            print("auto: claimText=" + str(claimText))

            if [txt for txt in claimText if "complete" in txt.lower()] != []:
                print("auto: complete=True")

                click(self.missionAcceptPos)
                print("auto: click=" + str(self.missionAutoPos))
                break;

            if count % 6 == 0:
                click((.5, .5))
                print("auto: click=" + str((.5, .5)))

                key("r")
                print("open: key=r") 

            count += 1
            time.sleep(2.5)
    
    def start(self):
        queue = [
            self.open, 
            self.accept, 
            self.close, 
            self.auto
        ]

        while True:
            for run in queue:
                run()
                time.sleep(0.4)

simpleBot = SimpleBot()
simpleBot.start()
