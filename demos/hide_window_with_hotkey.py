#!python3
# -*- coding: utf-8 -*-
# hide a window with hotkey Ctrl+1, show the hidden window with hotkey Ctrl+2
# run notepad.exe first and then press the hotkey for test
import os
import sys
import time
import ctypes
import subprocess
import psutil

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

def HotKeyFunc(stopEvent: 'Event', argv: list):
    args = [sys.executable, __file__] + argv
    cmd = ' '.join('"{}"'.format(arg) for arg in args)
    auto.Logger.WriteLine('call {}'.format(cmd))
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

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--main', action='store_true', help='exec main')
    parser.add_argument('--hide', action='store_true', help='hide window')
    parser.add_argument('--show', action='store_true', help='show window')
    args = parser.parse_args()

    if args.main:
        if args.hide:
            hide()
        elif args.show:
            show()
    else:
        subprocess.Popen('notepad')
        auto.GetConsoleWindow().SetActive()
        auto.Logger.ColorfullyWriteLine('Press <Color=Green>Ctr+1</Color> to hide\nPress <Color=Green>Ctr+2</Color> to show\n')
        auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): lambda e: HotKeyFunc(e, ['--main', '--hide']),
                            (auto.ModifierKey.Control, auto.Keys.VK_2): lambda e: HotKeyFunc(e, ['--main', '--show']),
                          },
                           (auto.ModifierKey.Control, auto.Keys.VK_9))

