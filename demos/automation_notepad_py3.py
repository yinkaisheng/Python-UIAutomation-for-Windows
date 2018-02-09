#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess
import ctypes

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation

text = """The uiautomation module

This module is for UIAutomation on Windows(Windows XP with SP3, Windows Vista and Windows 7/8/8.1/10).
It supports UIAutomation for the applications which implmented IUIAutomation, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt and Firefox.

Run 'automation.py -h' for help.

uiautomation is shared under the Apache Licene 2.0.
This means that the code can be freely copied and distributed, and costs nothing to use.

具体用法参考: http://www.cnblogs.com/Yinkaisheng/p/3444132.html
"""


def testNotepadCN():
    consoleWindow = automation.GetConsoleWindow()
    consoleWindow.SetActive()
    automation.Logger.ColorfulWriteLine('\nI will open <Color=Green>Notepad</Color> and <Color=Yellow>automate</Color> it. Please wait for a while.')
    time.sleep(2)
    automation.ShowDesktop()
    #打开notepad
    subprocess.Popen('notepad')
    #查找notepad， 如果name有中文，python2中要使用Unicode
    window = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad', RegexName = '.* - 记事本')
    #可以判断window是否存在，如果不判断，找不到window的话会抛出异常
    #if window.Exists(maxSearchSeconds = 3):
    if automation.WaitForExist(window, 3):
        automation.Logger.WriteLine("Notepad exists now")
    else:
        automation.Logger.WriteLine("Notepad does not exist after 3 seconds", automation.ConsoleColor.Yellow)
    screenWidth, screenHeight = automation.Win32API.GetScreenSize()
    window.MoveWindow(screenWidth // 4, screenHeight // 4, screenWidth // 2, screenHeight // 2)
    window.SetActive()
    #查找edit
    edit = automation.EditControl(searchFromControl = window)  #or edit = window.EditControl()
    edit.Click(waitTime = 0)
    #python2中要使用Unicode, 模拟按键
    edit.SetValue('hi你好')
    edit.SendKeys('{Ctrl}{End}{Enter}下面开始演示{! 4}{ENTER}', 0.2, 0)
    edit.SendKeys(text)
    edit.SendKeys('{Enter 3}0123456789{Enter}', waitTime = 0)
    edit.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{ENTER}', waitTime = 0)
    edit.SendKeys('abcdefghijklmnopqrstuvwxyz{ENTER}', waitTime = 0)
    edit.SendKeys('`~!@#$%^&*()-_=+{ENTER}', waitTime = 0)
    edit.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{ENTER}', waitTime = 0)
    edit.SendKeys('™®①②③④⑤⑥⑦⑧⑨⑩§№☆★○●◎◇◆□℃‰€■△▲※→←↑↓〓¤°＃＆＠＼＾＿―♂♀{ENTER}{CTRL}a')
    window.CaptureToImage('Notepad.png')
    edit.SendKeys('Image Notepad.png was captured, you will see it later.', 0.05)
    #查找菜单
    window.MenuItemControl(Name = '格式(O)').Click()
    window.MenuItemControl(Name = '字体(F)...').Click()
    windowFont = window.WindowControl(Name = '字体')
    windowFont.ComboBoxControl(AutomationId = '1140').Select('中文 GB2312')
    windowFont.ButtonControl(Name = '确定').Click()
    window.Close()
    if automation.WaitForDisappear(window, 3):
        automation.Logger.WriteLine("Notepad closed")
    else:
        automation.Logger.WriteLine("Notepad still exists after 3 seconds", automation.ConsoleColor.Yellow)

    # buttonNotSave = ButtonControl(searchFromControl = window, SubName = '不保存')
    # buttonNotSave.Click()
    # or send alt+n to not save and quit
    # automation.SendKeys('{Alt}n')
    # 使用另一种查找方法
    buttonNotSave = automation.FindControl(window,
        lambda control, depth: control.ControlType == automation.ControlType.ButtonControl and '不保存' in control.Name)
    buttonNotSave.Click()
    subprocess.Popen('Notepad.png', shell = True)
    time.sleep(2)
    consoleWindow.SetActive()
    automation.Logger.WriteLine('script exits', automation.ConsoleColor.Cyan)
    time.sleep(2)


def testNotepadEN():
    consoleWindow = automation.GetConsoleWindow()
    consoleWindow.SetActive()
    automation.Logger.ColorfulWriteLine('\nI will open <Color=Green>Notepad</Color> and <Color=Yellow>automate</Color> it. Please wait for a while.')
    time.sleep(2)
    automation.ShowDesktop()
    subprocess.Popen('notepad')
    #search notepad window, if searchFromControl is None, search from RootControl
    #searchDepth = 1 makes searching faster, only searches RootControl's children, not children's children
    window = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad', RegexName = '.* - Notepad')
    #if window.Exists(maxSearchSeconds = 3): #check before using it
    if automation.WaitForExist(window, 3):
        automation.Logger.WriteLine("Notepad exists now")
    else:
        automation.Logger.WriteLine("Notepad does not exist after 3 seconds", automation.ConsoleColor.Yellow)
    screenWidth, screenHeight = automation.Win32API.GetScreenSize()
    window.MoveWindow(screenWidth // 4, screenHeight // 4, screenWidth // 2, screenHeight // 2)
    window.SetActive()
    edit = automation.EditControl(searchFromControl = window)  #or edit = window.EditControl()
    edit.Click(waitTime = 0)
    edit.SetValue('hi你好')
    edit.SendKeys('{Ctrl}{End}{Enter}下面开始演示{! 4}{ENTER}', 0.2, 0)
    edit.SendKeys(text)
    edit.SendKeys('{Enter 3}0123456789{Enter}', waitTime = 0)
    edit.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{Enter}', waitTime = 0)
    edit.SendKeys('abcdefghijklmnopqrstuvwxyz{Enter}', waitTime = 0)
    edit.SendKeys('`~!@#$%^&*()-_=+{Enter}', waitTime = 0)
    edit.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{Enter}{Ctrl}a')
    window.CaptureToImage('Notepad.png')
    edit.SendKeys('Image Notepad.png was captured, you will see it later.', 0.05)
    #find menu
    window.MenuItemControl(Name = 'Format').Click()
    window.MenuItemControl(Name = 'Font...').Click()
    windowFont = window.WindowControl(Name = 'Font')
    windowFont.ComboBoxControl(AutomationId = '1140').Select('Western')
    windowFont.ButtonControl(Name = 'OK').Click()
    window.Close()
    if automation.WaitForDisappear(window, 3):
        automation.Logger.WriteLine("Notepad closed")
    else:
        automation.Logger.WriteLine("Notepad still exists after 3 seconds", automation.ConsoleColor.Yellow)
    # buttonNotSave = ButtonControl(searchFromControl = window, Name = 'Don\'t Save')
    # buttonNotSave.Click()
    # or send alt+n to not save and quit
    # automation.SendKeys('{Alt}n')
    # another way to find the button using lambda
    buttonNotSave = automation.FindControl(window,
        lambda control, depth: control.ControlType == automation.ControlType.ButtonControl and 'Don\'t Save' == control.Name)
    buttonNotSave.Click()
    subprocess.Popen('Notepad.png', shell = True)
    time.sleep(2)
    consoleWindow.SetActive()
    automation.Logger.WriteLine('script exits', automation.ConsoleColor.Cyan)
    time.sleep(2)

if __name__ == '__main__':
    uiLanguage = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    if uiLanguage == 2052:
        testNotepadCN()
    else:
        testNotepadEN()
