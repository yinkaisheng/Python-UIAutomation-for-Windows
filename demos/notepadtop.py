#!python3
# -*- coding:utf-8 -*-
import os
import sys
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
from uiautomation import uiautomation as automation

if __name__ == '__main__':
    isTop = 1
    print(sys.argv)

    if len(sys.argv) == 2:
        isTop = int(sys.argv[1])

    note = automation.WindowControl(searchDepth=1, ClassName='Notepad')
    if note.Exists(0, 0):
        note.SetTopmost(isTop)
    else:
        subprocess.Popen('notepad')
        note.Refind()
        note.SetTopmost(isTop)
        note.Move(0, 0)
        note.Resize(400, 300)
        edit = automation.EditControl(searchFromControl=note)
        edit.Click()
        automation.SendKeys('I\'m a topmost window!!!')

