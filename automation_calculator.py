#!python3
# -*- coding: utf-8 -*-
# works on windows XP, 7, 8 and 10
import time
import subprocess
import uiautomation as automation

def Calc(window, btns, expression):
    expression = ''.join(expression.split())
    if not expression.endswith('='):
        expression += '='
    for char in expression:
        automation.Logger.Write(char, writeToFile = False)
        btns[char].Click(waitTime = 0.05)
    window.SendKeys('{Ctrl}c', waitTime = 0)
    result = automation.Win32API.GetClipboardText()
    automation.Logger.WriteLine(result, automation.ConsoleColor.Cyan, writeToFile = False)
    time.sleep(1)

def CaclOnXP():
    chars = '0123456789.+-*/=()'
    calcWindow = automation.WindowControl(searchDepth = 1, ClassName = 'SciCalc')
    if not calcWindow.Exists(0, 0):
        subprocess.Popen('calc')
    calcWindow.SetTopmost()
    calcWindow.SendKeys('{Alt}vs', 0.5)
    clearBtn = calcWindow.ButtonControl(Name = 'CE')
    clearBtn.Click()
    char2Button = {}
    for key in chars:
        char2Button[key] = calcWindow.ButtonControl(Name = key)
    Calc(calcWindow, char2Button, '1234 * (4 + 5 + 6) - 78 / 90')
    Calc(calcWindow, char2Button, '2*3.14159*10')

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
    calcWindow = automation.WindowControl(searchDepth = 1, ClassName = 'CalcFrame')
    if not calcWindow.Exists(0, 0):
        subprocess.Popen('calc')
    calcWindow.SetTopmost()
    calcWindow.SendKeys('{Alt}2')
    clearBtn = calcWindow.ButtonControl(foundIndex= 8, Depth = 3)  #test foundIndex and Depth, the 8th button is clear
    if clearBtn.AutomationId == '82':
        clearBtn.Click()
    char2Button = {}
    for key in char2Id:
        char2Button[key] = calcWindow.ButtonControl(AutomationId = char2Id[key])
    Calc(calcWindow, char2Button, '1234 * (4 + 5 + 6) - 78 / 90')
    Calc(calcWindow, char2Button, '2*3.14159*10')

def CalcOnWindows10():
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
        '(' : 'openParanthesisButton',
        ')' : 'closeParanthesisButton',
    }
    calcWindow = automation.WindowControl(searchDepth = 1, ClassName = 'ApplicationFrameWindow', Name = 'Calculator')
    if not calcWindow.Exists(0, 0):
        subprocess.Popen('calc')
    calcWindow.SetTopmost()
    calcWindow.ButtonControl(AutomationId = 'NavButton').Click()
    calcWindow.ListItemControl(Name = 'Scientific Calculator').Click()
    calcWindow.ButtonControl(AutomationId = 'clearButton').Click()
    char2Button = {}
    for key in char2Id:
        char2Button[key] = calcWindow.ButtonControl(AutomationId = char2Id[key])
    Calc(calcWindow, char2Button, '1234 * (4 + 5 + 6) - 78 / 90')
    Calc(calcWindow, char2Button, '2*3.14159*10')

if __name__ == '__main__':
    import platform
    osVersion = int(platform.version().split('.')[0])
    if osVersion < 6:
        CaclOnXP()
    elif osVersion == 6:
        CalcOnWindows7And8()
    elif osVersion >= 10:
        CalcOnWindows10()
