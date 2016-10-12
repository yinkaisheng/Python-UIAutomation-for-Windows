#!python2
# -*- coding: utf-8 -*-
import os
import time
import subprocess
import ctypes
import automation

text = u'''The automation module

This module is for automation on Windows(Windows XP with SP3, Windows Vista and Windows 7/8/8.1/10).
It supports automation for the applications which implmented IUIAutomation, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt and Firefox.

Run 'automation.py -h' for help.

automation is shared under the MIT Licence.
This means that the code can be freely copied and distributed, and costs nothing to use.

具体用法参考: http://www.cnblogs.com/Yinkaisheng/p/3444132.html
'''

def testNotepadCN():
    automation.ShowDesktop()
    #打开notepad
    subprocess.Popen('notepad')
    #查找notepad， 如果name有中文，python2中要使用Unicode
    window = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad', SubName = u'无标题 - 记事本')
    #可以判断window是否存在，如果不判断，找不到window的话会抛出异常
    #if window.Exists(maxSearchSeconds = 3):
    screenWidth, screenHeight = automation.Win32API.GetScreenSize()
    window.MoveWindow(screenWidth // 4, screenHeight // 4, screenWidth // 2, screenHeight // 2)
    #查找edit
    edit = automation.EditControl(searchFromControl = window)
    edit.DoubleClick()
    time.sleep(1)
    #python2中要使用Unicode, 模拟按键
    edit.SetValue(u'hi你好')
    automation.SendKeys(u'{Ctrl}{End}{Enter}下面开始演示{! 4}{ENTER}', 0.2)
    automation.SendKeys(text)
    automation.SendKeys('{Enter 3}0123456789{Enter}')
    automation.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{Enter}')
    automation.SendKeys('abcdefghijklmnopqrstuvwxyz{Enter}')
    automation.SendKeys('`~!@#$%^&*()-_=+{ENTER}')
    automation.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{Enter}{Ctrl}a')
    time.sleep(1)
    window.CaptureToImage('Notepad.png')
    automation.SendKeys('Image Notepad.png was captured, you will see it later.', 0.05)
    time.sleep(1)
    #查找菜单
    menuItemFormat = automation.MenuItemControl(searchFromControl = window, Name = u'格式(O)')
    menuItemFont = automation.MenuItemControl(searchFromControl = window, Name = u'字体(F)...')
    windowFont = automation.WindowControl(searchFromControl = window, Name = u'字体')
    comboScript = automation.ComboBoxControl(searchFromControl = windowFont, AutomationId = '1140')
    buttonOK = automation.ButtonControl(searchFromControl = windowFont, Name = u'确定')
    menuItemFormat.Click()
    menuItemFont.Click() #or automation.SendKey(automation.Keys.VK_F)
    comboScript.Select(u'中文 GB2312')
    buttonOK.Click()
    time.sleep(1)
    window.Close()
    time.sleep(1)

    # buttonNotSave = ButtonControl(searchFromControl = window, SubName = u'不保存')
    # buttonNotSave.Click()
    # or send alt+n to not save and quit
    # automation.SendKeys('{Alt}n')
    # 使用另一种查找方法
    buttonNotSave = automation.FindControl(window,
        lambda control: control.ControlType == automation.ControlType.ButtonControl and u'不保存' in control.Name)
    buttonNotSave.Click()
    os.popen('Notepad.png')


def testNotepadEN():
    automation.ShowDesktop()
    subprocess.Popen('notepad')
    #search notepad window, if searchFromControl is None, search from RootControl
    #searchDepth = 1 makes searching faster, only searches RootControl's children, not children's children
    window = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad', SubName = 'Untitled - Notepad')
    #if window.Exists(maxSearchSeconds = 3): #check before using it
        #pass
    screenWidth, screenHeight = automation.Win32API.GetScreenSize()
    window.MoveWindow(screenWidth // 4, screenHeight // 4, screenWidth // 2, screenHeight // 2)
    edit = automation.EditControl(searchFromControl = window)
    window.SetActive()
    time.sleep(1)
    edit.DoubleClick()
	#use unicode string
    edit.SetValue(u'hi你好')
    automation.SendKeys(u'{Ctrl}{End}{Enter}下面开始演示{! 4}{ENTER}', 0.2)
    automation.SendKeys(text)
    automation.SendKeys('{ENTER 3}0123456789{ENTER}')
    automation.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{ENTER}')
    automation.SendKeys('abcdefghijklmnopqrstuvwxyz{ENTER}')
    automation.SendKeys('`~!@#$%^&*()-_=+{ENTER}')
    automation.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{ENTER}{CTRL}a')
    time.sleep(1)
    window.CaptureToImage('Notepad.png')
    automation.SendKeys('Image Notepad.png was captured, you will see it later.', 0.05)
    time.sleep(1)
    menuItemFormat = automation.MenuItemControl(searchFromControl = window, Name = 'Format')
    menuItemFont = automation.enuItemControl(searchFromControl = window, Name = 'Font...')
    windowFont = automation.WindowControl(searchFromControl = window, Name = 'Font')
    comboScript = automation.ComboBoxControl(searchFromControl = windowFont, AutomationId = '1140')
    buttonOK = automation.ButtonControl(searchFromControl = windowFont, Name = 'OK')
    menuItemFormat.Click()
    menuItemFont.Click() #or automation.SendKey(automation.Keys.VK_F)
    comboScript.Select('Western')
    buttonOK.Click()
    time.sleep(1)
    window.Close()
    time.sleep(1)
    # buttonNotSave = ButtonControl(searchFromControl = window, Name = 'Don\'t Save')
    # buttonNotSave.Click()
    # or send alt+n to not save and quit
    # automation.SendKeys('{Alt}n')
    # another way to find the button using lambda
    buttonNotSave = automation.FindControl(window,
        lambda control: control.ControlType == automation.ControlType.ButtonControl and 'Don\'t Save' == control.Name)
    buttonNotSave.Click()
    os.popen('Notepad.png')

if __name__ == '__main__':
    uiLanguage = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    if uiLanguage == 2052:
        testNotepadCN()
    else:
        testNotepadEN()
