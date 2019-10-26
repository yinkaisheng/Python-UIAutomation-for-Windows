#!python3
# -*- coding: utf-8 -*-
# works on windows XP, 7, 8, 8.1 and 10
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

auto.uiautomation.DEBUG_EXIST_DISAPPEAR = True  # set it to False and try again, default is False
auto.uiautomation.DEBUG_SEARCH_TIME = True  # set it to False and try again, default is False
auto.uiautomation.TIME_OUT_SECOND = 10  # global time out


def Calc(window, btns, expression):
    expression = ''.join(expression.split())
    if not expression.endswith('='):
        expression += '='
    for char in expression:
        auto.Logger.Write(char, writeToFile = False)
        btns[char].Click(waitTime=0.05)
    time.sleep(0.1)
    window.SendKeys('{Ctrl}c', waitTime = 0.1)
    result = auto.GetClipboardText()
    auto.Logger.WriteLine(result, auto.ConsoleColor.Cyan, writeToFile = False)
    time.sleep(1)


def CalcOnXP():
    chars = '0123456789.+-*/=()'
    #Desc is not a valid search property, but it can be used for debug printing
    calcWindow = auto.WindowControl(searchDepth=1, ClassName='SciCalc', Desc='Calculator Window')
    if not calcWindow.Exists(0, 0):
        subprocess.Popen('calc')
    calcWindow.SetActive()
    calcWindow.SendKeys('{Alt}vs', 0.5)
    clearBtn = calcWindow.ButtonControl(Name = 'CE')  #
    clearBtn.Click()
    if 0:
        char2Button = {key: calcWindow.ButtonControl(Name=key, Desc='Button ' + key) for key in chars}
    else:
        #Run faster because it only walk calc window once
        char2Button = {}
        for c, d in auto.WalkControl(calcWindow, maxDepth=1):
            if c.Name in chars:
                char2Button[c.Name] = c
    Calc(calcWindow, char2Button, '1234 * (4 + 5 + 6) - 78 / 90.8')
    Calc(calcWindow, char2Button, '3*3+4*4')
    Calc(calcWindow, char2Button, '2*3.14159*10')
    calcWindow.Disappears(1)
    calcWindow.GetWindowPattern().Close()
    calcWindow.Exists(1)


def CalcOnWindows7And8():
    char2Id = {
        '0' : '130',
        '1' : '131',
        '2' : '132',
        '3' : '133',
        '4' : '134',
        '5' : '135',
        '6' : '136',
        '7' : '137',
        '8' : '138',
        '9' : '139',
        '.' : '84',
        '+' : '93',
        '-' : '94',
        '*' : '92',
        '/' : '91',
        '=' : '121',
        '(' : '128',
        ')' : '129',
    }
    #Desc is not a valid search property, but it can be used for debug printing
    calcWindow = auto.WindowControl(searchDepth = 1, ClassName = 'CalcFrame', Desc='Calculator Window')
    if not calcWindow.Exists(0, 0):
        subprocess.Popen('calc')
    calcWindow.SetActive()
    calcWindow.SendKeys('{Alt}2')
    clearBtn = calcWindow.ButtonControl(foundIndex= 8, Depth = 3)  #test foundIndex and Depth, the 8th button is clear
    if clearBtn.Exists() and clearBtn.AutomationId == '82':
        clearBtn.Click()
    if 0:
        char2Button = {key: calcWindow.ButtonControl(AutomationId=char2Id[key], Desc='Button ' + key) for key in char2Id}
    else:
        #Run faster because it only walk calc window once
        id2char = {v: k for k, v in char2Id.items()}
        char2Button = {}
        for c, d in auto.WalkControl(calcWindow):
            if c.AutomationId in id2char:
                char2Button[id2char[c.AutomationId]] = c
    Calc(calcWindow, char2Button, '1234 * (4 + 5 + 6) - 78 / 90.8')
    Calc(calcWindow, char2Button, '3*3+4*4')
    Calc(calcWindow, char2Button, '2*3.14159*10')
    calcWindow.Disappears(1)
    calcWindow.GetWindowPattern().Close()
    calcWindow.Exists(1)


def CalcOnWindows10():
    """works on Windows 10.0.18362"""
    char2Id = {
        '0' : 'num0Button',
        '1' : 'num1Button',
        '2' : 'num2Button',
        '3' : 'num3Button',
        '4' : 'num4Button',
        '5' : 'num5Button',
        '6' : 'num6Button',
        '7' : 'num7Button',
        '8' : 'num8Button',
        '9' : 'num9Button',
        '.' : 'decimalSeparatorButton',
        '+' : 'plusButton',
        '-' : 'minusButton',
        '*' : 'multiplyButton',
        '/' : 'divideButton',
        '=' : 'equalButton',
        '(' : 'openParenthesisButton',
        ')' : 'closeParenthesisButton',
    }
    #Desc is not a valid search property, but it can be used for debug printing
    calcWindow = auto.WindowControl(searchDepth = 1, ClassName = 'ApplicationFrameWindow',
        Compare = lambda c,d : c.Name == 'Calculator' or c.Name == '计算器', Desc='Calculator Window')
    if not calcWindow.Exists(0, 0):
        subprocess.Popen('calc')
    calcWindow.SetActive()
    calcWindow.ButtonControl(AutomationId = 'TogglePaneButton').Click()
    calcWindow.ListItemControl(AutomationId = 'Scientific').Click()
    calcWindow.ButtonControl(AutomationId='clearButton').Click()
    if 1:
        char2Button = {key: calcWindow.ButtonControl(AutomationId=char2Id[key], Desc='Button ' + key) for key in char2Id}
    else:
        #Run faster because it only walk calc window once
        id2char = {v: k for k, v in char2Id.items()}
        char2Button = {}
        for c, d in auto.WalkControl(calcWindow, maxDepth=3):
            if c.AutomationId in id2char:
                char2Button[id2char[c.AutomationId]] = c
    Calc(calcWindow, char2Button, '1234 * (4 + 5 + 6) - 78 / 90.8')
    Calc(calcWindow, char2Button, '3*3+4*4')
    Calc(calcWindow, char2Button, '2*3.14159*10')
    calcWindow.CaptureToImage('calc.png', 7, 0, -14, -7)  # on windows 10, 7 pixels of windows border are transparent
    calcWindow.Disappears(1)
    calcWindow.GetWindowPattern().Close()
    calcWindow.Exists(1)

if __name__ == '__main__':
    osVersion = os.sys.getwindowsversion().major
    if osVersion < 6:
        CalcOnXP()
    elif osVersion == 6:
        CalcOnWindows7And8()
    elif osVersion >= 10:
        CalcOnWindows10()

    auto.Logger.Write('\nPress any key to exit', auto.ConsoleColor.Cyan)
    import msvcrt
    while not msvcrt.kbhit():
        pass
