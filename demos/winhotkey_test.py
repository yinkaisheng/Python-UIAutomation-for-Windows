#!python3
# -*- coding:utf-8 -*-
import os
import sys
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def test1(stopEvent):
    c = auto.GetRootControl()
    n = 0
    while True:
        if stopEvent.is_set():
            break
        print(n)
        n += 1
        stopEvent.wait(1)
    print('test1 exits')


def test2(stopEvent):
    p = subprocess.Popen('python.exe automation_notepad_py3.py')
    auto.Logger.WriteLine('call python.exe automation_notepad_py3.py', auto.ConsoleColor.DarkGreen)
    while True:
        if None != p.poll():
            break
        if stopEvent.is_set():
            childProcess = []
            for pid, pname in auto.Win32API.EnumProcess():
                ppid = auto.Win32API.GetParentProcessId(pid)
                if ppid == p.pid or pid == p.pid:
                    cmd = auto.Win32API.GetProcessCommandLine(pid)
                    childProcess.append((pid, pname, cmd))
            for pid, pname, cmd in childProcess:
                auto.Logger.WriteLine('kill process: {}, {}, "{}"'.format(pid, pname, cmd), auto.ConsoleColor.Yellow)
                auto.Win32API.TerminateProcess(pid)
            break
        stopEvent.wait(0.05)
    auto.Logger.WriteLine('test2 exits', auto.ConsoleColor.DarkGreen)


def main():
    auto.Logger.WriteLine('Press Ctrl+1 to start, press ctrl+4 to stop')
    auto.RunWithHotKey({(auto.ModifierKey.MOD_CONTROL, auto.Keys.VK_1): test1,
                              (auto.ModifierKey.MOD_CONTROL, auto.Keys.VK_2): test2},
                             (auto.ModifierKey.MOD_CONTROL, auto.Keys.VK_4))

if __name__ == '__main__':
    main()
