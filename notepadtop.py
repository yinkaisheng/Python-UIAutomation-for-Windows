#!python3
# -*- coding:utf-8 -*-
import sys
import subprocess
import uiautomation as automation

if __name__ == '__main__':
    isTop = 1
    print(sys.argv)
    if len(sys.argv) == 2:
        isTop = int(sys.argv[1])
    note = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad')
    if note.Exists(0, 0):
        note.SetTopmost(isTop)
    else:
        subprocess.Popen('notepad')
        note.Refind()
        note.SetTopmost(isTop)
        note.Move(0, 0)
        note.Resize(400, 300)
        edit = automation.EditControl(searchFromControl= note)
        edit.Click()
        automation.SendKeys('I\'m a topmost window!!!')

