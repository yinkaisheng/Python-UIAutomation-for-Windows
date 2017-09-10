#!python3
# -*- coding:utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation

if __name__ == '__main__':
    automation.ShowDesktop()
    note = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad')
    if note.Exists(0, 0):
        note.SetActive()
    else:
        subprocess.Popen('notepad')
        note.Refind()
    note.SetTopmost()
    note.Move(400, 0)
    note.Resize(400, 300)
    edit = note.EditControl()
    edit.SendKeys('{Ctrl}{End}{Enter 2}I\'m a topmost window!!!\nI cover other windows.')
    subprocess.Popen('mmc.exe devmgmt.msc')
    mmcWindow = automation.WindowControl(searchDepth = 1, ClassName = 'MMCMainFrame')
    mmcWindow.Move(100, 100)
    mmcWindow.Resize(400, 300)
    time.sleep(1)
    automation.Win32API.MouseDragDrop(160, 110, 900, 110, 0.2)
    automation.Win32API.MouseDragDrop(900, 110, 160, 110, 0.2)
    #mmcWindow.Maximize(waitTime= 1)
    mmcWindow.SendKeys('{Alt}f', waitTime= 1)
    mmcWindow.SendKey(automation.Keys.VK_X)
    #edit.SendKeys('{Ctrl}z')


