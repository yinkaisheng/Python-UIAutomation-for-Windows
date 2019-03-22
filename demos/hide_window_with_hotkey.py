#!python3
# -*- coding: utf-8 -*-
# hide a window with hotkey Ctrl+1, show the hidden window with hotkey Ctrl+2
# run notepad.exe first and then press the hotkey for test
import os
import sys
import time
import ctypes
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

WindowsWantToHide = ('Warcraft III', 'Valve001', 'Counter-Strike', 'Notepad')


def hide():
    root = auto.GetRootControl()
    for window in root.GetChildren():
        if window.ClassName in WindowsWantToHide:
            auto.Logger.WriteLine('hide window, handle {}'.format(window.NativeWindowHandle))
            window.Hide()
            fin = open('hide_windows.txt', 'wt')
            fin.write(str(window.NativeWindowHandle) + '\n')
            fin.close()


def show():
    fout = open('hide_windows.txt')
    lines = fout.readlines()
    fout.close()
    for line in lines:
        handle = int(line)
        window = auto.ControlFromHandle(handle)
        if window:
            auto.Logger.WriteLine('show window: {}'.format(handle))
            window.Show()

def HideWindowFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'python.exe {} hide {}'.format(scriptName, ' '.join(sys.argv[1:]))
    auto.Logger.ColorfullyWriteLine('HideWindowFunc call <Color=Green>{}</Color>'.format(cmd))
    p = subprocess.Popen(cmd)
    while True:
        if None != p.poll():
            break
        if stopEvent.is_set():
            childProcesses = [pro for pro in psutil.process_iter() if pro.ppid == p.pid or pro.pid == p.pid]
            for pro in childProcesses:
                auto.Logger.WriteLine('kill process: {}, {}'.format(pro.pid, pro.cmdline()), auto.ConsoleColor.Yellow)
                p.kill()
            break
        stopEvent.wait(0.01)
    auto.Logger.WriteLine('HideWindowFunc exit')

def ShowWindowFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'python.exe {} show {}'.format(scriptName, ' '.join(sys.argv[1:]))
    auto.Logger.ColorfullyWriteLine('ShowWindowFunc call <Color=Green>{}</Color>'.format(cmd))
    p = subprocess.Popen(cmd)
    while True:
        if None != p.poll():
            break
        if stopEvent.is_set():
            childProcesses = [pro for pro in psutil.process_iter() if pro.ppid == p.pid or pro.pid == p.pid]
            for pro in childProcesses:
                auto.Logger.WriteLine('kill process: {}, {}'.format(pro.pid, pro.cmdline()), auto.ConsoleColor.Yellow)
                p.kill()
            break
        stopEvent.wait(0.01)
    auto.Logger.WriteLine('ShowWindowFunc exit')

if __name__ == '__main__':
    if 'hide' in sys.argv[1:]:
        hide()
    elif 'show' in sys.argv[1:]:
        show()
    else:
        subprocess.Popen('notepad')
        auto.GetConsoleWindow().SetActive()
        auto.Logger.ColorfullyWriteLine('Run Notepad\nPress <Color=Green>Ctr+1</Color> to hide\nPress <Color=Green>Ctr+2</Color> to show\n')
        auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): HideWindowFunc, (auto.ModifierKey.Control, auto.Keys.VK_2): ShowWindowFunc}, (auto.ModifierKey.Control, auto.Keys.VK_4)
                         )
