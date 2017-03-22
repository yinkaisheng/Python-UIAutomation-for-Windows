#!python3
# -*- coding: utf-8 -*-
# test only on windows 7
import time
import subprocess
import uiautomation as automation

def main():
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
    def calc(expression):
        expression = ''.join(expression.split())
        if not expression.endswith('='):
            expression += '='
        for char in expression:
            automation.Logger.Write(char, writeToFile = False)
            char2Button[char].Click(waitTime = 0.05)
        calcWindow.SendKeys('{Ctrl}c', waitTime = 0)
        result = automation.Win32API.GetClipboardText()
        automation.Logger.WriteLine(result, automation.ConsoleColor.Cyan, writeToFile = False)
        time.sleep(1)
    calc('2*3.14159*10')
    calc('1234 * (4 + 5 + 6) - 78 / 90')


if __name__ == '__main__':
    main()
