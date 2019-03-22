#!python3
# -*- coding:utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

def main():
    auto.ShowDesktop()
    subprocess.Popen('notepad')
    time.sleep(1)
    note = auto.WindowControl(searchDepth=1, ClassName='Notepad')
    note.SetActive()
    note.SetTopmost()
    transformNote = note.GetTransformPattern()
    transformNote.Move(400, 0)
    transformNote.Resize(400, 300)
    edit = note.EditControl()
    edit.SendKeys('{Ctrl}{End}{Enter 2}I\'m a topmost window!!!\nI cover other windows.')
    subprocess.Popen('mmc.exe devmgmt.msc')
    mmcWindow = auto.WindowControl(searchDepth=1, ClassName='MMCMainFrame')
    mmcWindow.SetActive()
    transformMmc = mmcWindow.GetTransformPattern()
    transformMmc.Move(100, 100)
    transformMmc.Resize(400, 300)
    time.sleep(1)
    auto.DragDrop(160, 110, 900, 110, 0.2)
    auto.DragDrop(900, 110, 160, 110, 0.2)
    mmcWindow.SendKeys('{Alt}f', waitTime=1)
    mmcWindow.SendKey(auto.Keys.VK_X)
    auto.GetConsoleWindow().SetActive()

if __name__ == '__main__':
    main()

