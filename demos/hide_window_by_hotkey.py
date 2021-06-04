#!python3
# -*- coding: utf-8 -*-
# hide windows with hotkey Ctrl+1, show the hidden windows with hotkey Ctrl+2
import os
import sys
import time
import subprocess
from typing import List
from threading import Event

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

WindowsWantToHide = ('Warcraft III', 'Valve001', 'Counter-Strike', 'Notepad')


def hide(stopEvent: Event, handles: List[int]):
    with auto.UIAutomationInitializerInThread():
        for handle in handles:
            win = auto.ControlFromHandle(handle)
            win.Hide(0)


def show(stopEvent: Event, handle: List[int]):
    with auto.UIAutomationInitializerInThread():
        for handle in handles:
            win = auto.ControlFromHandle(handle)
            win.Show(0)
            if auto.IsIconic(handle):
                win.ShowWindow(auto.SW.Restore, 0)


if __name__ == '__main__':
    for i in range(2):
        subprocess.Popen('notepad.exe')
        time.sleep(1)
        notepad = auto.WindowControl(searchDepth=1, ClassName='Notepad')
        notepad.MoveWindow(i * 400, 0, 400, 300)
        notepad.SendKeys('notepad {}'.format(i + 1))
    auto.SetConsoleTitle('Hide: Ctrl+1, Show: Ctrl+2, Exit: Ctrl+D')
    cmdWindow = auto.GetConsoleWindow()
    if cmdWindow:
        cmdWindow.GetTransformPattern().Move(0, 300)
    auto.Logger.ColorfullyWriteLine('Press <Color=Green>Ctr+1</Color> to hide the windows\nPress <Color=Green>Ctr+2</Color> to show the windows\n')
    handles = [win.NativeWindowHandle for win in auto.GetRootControl().GetChildren() if win.ClassName in WindowsWantToHide]
    auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): lambda event: hide(event, handles),
                      (auto.ModifierKey.Control, auto.Keys.VK_2): lambda event: show(event, handles),
                      }, waitHotKeyReleased=False)

