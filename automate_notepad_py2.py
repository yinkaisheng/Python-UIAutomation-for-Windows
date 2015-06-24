#!python2
# -*- coding:utf-8 -*-
import time
import subprocess
import ctypes
import automation

text = '''# The automation module

This module is for automation on Windows{(}Windows XP with SP3, Windows Vista, Windows 7 and Windows 8/8.1{)}.
It supports automation for the applications which implmented IUIAutomation, such as MFC, Windows Forms, WPF, Windows 8 Metro App, Qt and Firefox.

Run 'automation.py -h' for help.

automation is shared under the GNU Public Licence (GPLv2, v3, and the AGPLv3).
This means that the code can be freely copied and distributed, and costs nothing to use in other GPL applications.

The details: http://www.cnblogs.com/Yinkaisheng/p/3444132.html
'''

def testNotepadCN():
    automation.ShowDesktop()
    #打开notepad
    subprocess.Popen('notepad')
    #查找notepad， 如果name有中文，python2中要使用Unicode
    window = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad', SubName = u'无标题 - 记事本')
    #可以判断window是否存在，如果不判断，找不到window的话会抛出异常
    #if window.Exists(maxSearchSeconds = 3):
    #查找edit
    edit = automation.EditControl(searchFromControl = window)
    edit.DoubleClick()
    time.sleep(1)
    #python2中要使用Unicode, 模拟按键
    edit.SetValue(u'hi你好')
    automation.Win32API.SendKeys(u'{Ctrl}{End}{Enter}下面开始演示{! 4}{ENTER}', 0.2)
    automation.Win32API.SendKeys(text)
    automation.Win32API.SendKeys('{ENTER 3}0123456789{ENTER}')
    automation.Win32API.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{ENTER}')
    automation.Win32API.SendKeys('abcdefghijklmnopqrstuvwxyz{ENTER}')
    automation.Win32API.SendKeys('`~!@#$%^&*{(}{)}_=+{ENTER}')
    automation.Win32API.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{ENTER}{CTRL}a')
    time.sleep(1)
    #查找菜单
    menuItemFormat = automation.MenuItemControl(searchFromControl = window, Name = u'格式(O)')
    menuItemFont = automation.MenuItemControl(searchFromControl = window, Name = u'字体(F)...')
    windowFont = automation.WindowControl(searchFromControl = window, Name = u'字体')
    comboScript = automation.ComboBoxControl(searchFromControl = windowFont, AutomationId = '1140')
    buttonOK = automation.ButtonControl(searchFromControl = windowFont, Name = u'确定')
    menuItemFormat.Click()
    menuItemFont.Click()
    comboScript.Select(u'中文 GB2312')
    buttonOK.Click()
    time.sleep(1)
    window.Close()
    time.sleep(1)
    
    # buttonNotSave = ButtonControl(searchFromControl = window, SubName = u'不保存')
    # buttonNotSave.Click()
    # or send alt+n to not save and quit
    # Win32API.SendKeys('{ALT}n')
    # 使用另一种查找方法
    buttonNotSave = automation.FindControl(window,
        lambda control: control.ControlType == automation.ControlType.ButtonControl and u'不保存' in control.Name)
    buttonNotSave.Click()


def testNotepadEN():
    automation.ShowDesktop()
    subprocess.Popen('notepad')
    window = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad', SubName = 'Untitled - Notepad')
    edit = automation.EditControl(searchFromControl = window)
    window.SetActive()
    time.sleep(1)
    edit.DoubleClick()
    edit.SetValue(u'hi你好')
    automation.Win32API.SendKeys(u'{Ctrl}{End}{Enter}下面开始演示{! 4}{ENTER}', 0.2)
    automation.Win32API.SendKeys(text)
    automation.Win32API.SendKeys('{ENTER 3}0123456789{ENTER}')
    automation.Win32API.SendKeys('ABCDEFGHIJKLMNOPQRSTUVWXYZ{ENTER}')
    automation.Win32API.SendKeys('abcdefghijklmnopqrstuvwxyz{ENTER}')
    automation.Win32API.SendKeys('`~!@#$%^&*()-_=+{ENTER}')
    automation.Win32API.SendKeys('[]{{}{}}\\|;:\'\",<.>/?{ENTER}{CTRL}a')
    time.sleep(1)
    menuItemFormat = automation.MenuItemControl(searchFromControl = window, Name = 'Format')
    menuItemFont = automation.enuItemControl(searchFromControl = window, Name = 'Font...')
    windowFont = automation.WindowControl(searchFromControl = window, Name = 'Font')
    comboScript = automation.ComboBoxControl(searchFromControl = windowFont, AutomationId = '1140')
    buttonOK = automation.ButtonControl(searchFromControl = windowFont, Name = 'OK')
    menuItemFormat.Click()
    menuItemFont.Click()
    comboScript.Select('Western')
    buttonOK.Click()
    time.sleep(1)
    window.Close()
    time.sleep(1)
    # buttonNotSave = ButtonControl(searchFromControl = window, Name = 'Don\'t Save')
    # buttonNotSave.Click()
    # or send alt+n to not save and quit
    # Win32API.SendKeys('{ALT}n')
    # 使用另一种查找方法
    buttonNotSave = automation.FindControl(window,
        lambda control: control.ControlType == automation.ControlType.ButtonControl and 'Don\'t Save' == control.Name)
    buttonNotSave.Click()

if __name__ == '__main__':
    uiLanguage = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    if uiLanguage == 2052:
        testNotepadCN()
    else:
        testNotepadEN()
