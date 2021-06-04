#!python3
# -*- coding: utf-8 -*-
# hide windows with hotkey Ctrl+1, show the hidden windows with hotkey Ctrl+2
import os
import sys
import time
import subprocess
from threading import Event

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def capture(stopEvent: Event):
    with auto.UIAutomationInitializerInThread(debug=True):
        control = auto.ControlFromCursor()
        control.CaptureToImage('control.png')
        subprocess.Popen('control.png', shell=True)


if __name__ == '__main__':
    auto.SetConsoleTitle('Capture: Ctrl+1, Exit: Ctrl+D')
    auto.Logger.ColorfullyWriteLine('Press <Color=Green>Ctr+1</Color> to capture a control image under mouse corsor')
    auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): capture,
                      })


