#!python3
# -*- coding:utf-8 -*-
import os
import sys
import subprocess
import psutil

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def test1(stopEvent):
    n = 0
    while True:
        if stopEvent.is_set():
            break
        print(n)
        n += 1
        stopEvent.wait(1)
    print('test1 exits')


def test2(stopEvent):
    p = subprocess.Popen('python.exe automation_notepad.py')
    auto.Logger.WriteLine('call python.exe automation_notepad.py', auto.ConsoleColor.DarkGreen)
    while True:
        if None != p.poll():
            break
        if stopEvent.is_set():
            childProcesses = [pro for pro in psutil.process_iter() if pro.ppid == p.pid or pro.pid == p.pid]
            for pro in childProcesses:
                auto.Logger.WriteLine('kill process: {}, {}'.format(pro.pid, pro.cmdline()), auto.ConsoleColor.Yellow)
                p.kill()
            break
        stopEvent.wait(0.05)
    auto.Logger.WriteLine('test2 exits', auto.ConsoleColor.DarkGreen)


def main():
    if auto.IsNT6orHigher:
        oriTitle = auto.GetConsoleOriginalTitle()
    else:
        oriTitle = auto.GetConsoleTitle()
    auto.SetConsoleTitle(auto.GetConsoleTitle() + ' | Ctrl+1,Ctrl+2, stop Ctrl+4')
    auto.Logger.ColorfullyWriteLine('Press <Color=Cyan>Ctrl+1 or Ctrl+2</Color> to start, press <Color=Cyan>ctrl+4</Color> to stop')
    auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): test1,
                              (auto.ModifierKey.Control, auto.Keys.VK_2): test2},
                             (auto.ModifierKey.Control, auto.Keys.VK_4))
    auto.SetConsoleTitle(oriTitle)

if __name__ == '__main__':
    main()
