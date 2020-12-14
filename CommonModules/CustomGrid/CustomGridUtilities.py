#-------------------------------------------------------------------------------
# Name:        MyCommon
# Purpose:     General purpose utilities.
#
#
# Author:      Douglas Parent
#
# Created:     12/20/2011
# Copyright:   (c) Douglas Parent 2011
# License:     You are free to use and modify this software for your own use.
#-------------------------------------------------------------------------------
# General purpose functions

import win32process
import win32api
import win32con
import win32com.client
import win32clipboard
import sys
from threading import Thread
import pythoncom
import time

########################
# on the clipboard methods, wait 50 ms before using the clipboard
# many times we use the clipboard after having issued a clipboard action
# So need to leave some time for the clipboard to finish its previous action
def setClipboard(text):
    time.sleep(0.05)
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_TEXT, text)
    except:
        pass
    finally:
        win32clipboard.CloseClipboard()

def getClipboard():
    time.sleep(0.05)
    text = ""
    try:
        win32clipboard.OpenClipboard()
        text = win32clipboard.GetClipboardData(win32con.CF_TEXT)
    except:
        pass
    finally:
        win32clipboard.CloseClipboard()

    # Remove trailing null character
    # print repr(text)
    if len(text) > 0:
        if '\0' in text:
            text = text[:text.index('\0')]

    return text

def emptyClipboard():
    time.sleep(0.05)
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
    except:
        pass
    finally:
        win32clipboard.CloseClipboard()

def getHwndProcessName(hwnd):
    retTuple = win32process.GetWindowThreadProcessId(hwnd)
    if retTuple:
        pid = retTuple[1]
        pHandle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION | win32con.PROCESS_VM_READ, 0, int(pid))
        if pHandle:
            return win32process.GetModuleFileNameEx(pHandle, 0)

# The following method stores the found hwnd in this global variable
foundProcessHwnd = None
def checkHwndForProcessName(hwnd, name):
    global foundProcessHwnd
    foundProcessHwnd = None
    filename = getHwndProcessName(hwnd)
    print filename
    if filename.find(name) > -1:
        foundProcessHwnd = hwnd
        # stop enumeration
        return 0
    else:
        # continue enumeration
        return 1

##########################
# Message Box methods
##########################
# Full details:  http://msdn.microsoft.com/en-us/library/ms645505%28VS.85%29.aspx
# Button Types
MESSAGEBOX_OK = 0x0
MESSAGEBOX_OK_CANCEL = 0x1
MESSAGEBOX_ABORT_RETRY_IGNORE = 0x2
MESSAGEBOX_YES_NO_CANCEL = 0x3
MESSAGEBOX_YES_NO = 0x4
MESSAGEBOX_RETRY_CANCEL = 0x5
MESSAGEBOX_CANCEL_TRY_CONTINUE = 0x6

# Icon Types
ICON_STOP = 0x10
ICON_QUESTION = 0x20
ICON_EXCLAMATION = 0x30
ICON_INFORMATION = 0x40

# Other type values
DEFAULT_SECOND_BUTTON = 0x100
DEFAULT_THIRD_BUTTON = 0x200
SYSTEM_MODAL = 0x1000
RIGHT_JUSTIFIED = 0x80000
RIGHT_TO_LEFT_READING_ORDER = 0x100000

# Return values
MESSAGEBOX_RETURN_TIMEOUT = -1
MESSAGEBOX_RETURN_OK = 1
MESSAGEBOX_RETURN_CANCEL = 2
MESSAGEBOX_RETURN_ABORT = 3
MESSAGEBOX_RETURN_RETRY = 4
MESSAGEBOX_RETURN_IGNORE = 5
MESSAGEBOX_RETURN_YES = 6
MESSAGEBOX_RETURN_NO = 7
MESSAGEBOX_RETURN_TRY_AGAIN = 10
MESSAGEBOX_RETURN_CONTINUE = 11

MESSAGEBOX_APP_MODAL = 0
MESSAGEBOX_SYS_MODAL = 0x1000
MESSAGEBOX_TASK_MODAL = 0x2000
MESSAGEBOX_SERVICE_NOTIFICATION = 0x200000

# do not run any of the message box from a natlink thread.
# it may cause the entire NaturallySpeaking process to abruptly terminate.
# Instead, call one of the methods on the MessageBoxThread class.
class MessageBoxThread(Thread):
    global MESSAGEBOX_OK, MESSAGEBOX_OK_CANCEL, MESSAGEBOX_YES_NO, MESSAGEBOX_YES_NO_CANCEL
    global ICON_QUESTION, ICON_INFORMATION

    returnValue = True

    # Defines some common message box styles
    # Pass these into the object constructor
    STYLE_OK = MESSAGEBOX_OK + ICON_INFORMATION
    STYLE_OK_CANCEL = MESSAGEBOX_OK_CANCEL + ICON_QUESTION
    STYLE_YES_NO = MESSAGEBOX_YES_NO + ICON_QUESTION
    STYLE_YES_NO_CANCEL = MESSAGEBOX_YES_NO_CANCEL + ICON_QUESTION

    def __init__(self, prompt, title, style = MESSAGEBOX_OK + ICON_INFORMATION):
        Thread.__init__(self)
        self.prompt = prompt
        self.title = title
        self.style = style

    def run(self):
        sys.conit_flags = 0
        pythoncom.CoInitializeEx(0)
        shell = win32com.client.Dispatch("WScript.Shell")
        self.returnValue = shell.Popup(self.prompt, 0, self.title, self.style)
        """
        vbhost = win32com.client.Dispatch("ScriptControl")
        vbhost.language = "vbscript"
        vbhost.AllowUI = "True"
        returnValue = vbhost.Eval('InputBox("Hi there")')
        vbhost.Eval('MsgBox("' + returnValue + '")')
        print "test"
        print returnValue

        """
        #for id in range(1, 5000):
        #    print id

        pythoncom.CoUninitialize()
        return

################################
# Non-thread version of Messagebox

# Returns one of the return values
def messageBox(prompt, title, style):
    shell = win32com.client.Dispatch("WScript.Shell")
    return shell.Popup(prompt, 0, title, style)

# Always returns true
def messageBoxOK(prompt, title):
    return messageBox(prompt, title, MESSAGEBOX_OK + ICON_INFORMATION)

# Returns True if the default button is clicked, False if the secondary button is clicked.
def messageBoxOKCancel(prompt, title):
    return messageBox(prompt, title, MESSAGEBOX_OK_CANCEL + ICON_QUESTION)

# Returns True if the default button is clicked, False if the secondary button is clicked.
def messageBoxYesNo(prompt, title):
    return messageBox(prompt, title, MESSAGEBOX_YES_NO + ICON_QUESTION)

# Returns True if the default button is clicked, False if the secondary button is clicked.
def messageBoxYesNoCancel(prompt, title):
    return messageBox(prompt, title, MESSAGEBOX_YES_NO_CANCEL + ICON_QUESTION)

def unload():
    global foundProcessHwnd
    foundProcessHwnd = None
