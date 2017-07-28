#!python3
# -*- coding:utf-8 -*-
import os
import time
import subprocess
import uiautomation as automation

def ReleaseAllKey():
    for key, value in automation.Keys.__dict__.items():
        if isinstance(value, int) and key.startswith('VK'):
            if automation.Win32API.IsKeyPressed(value):
                automation.Logger.WriteLine('{:<30} is pressed, release it'.format(key))
                automation.Win32API.ReleaseKey(value)

def test1(stopEvent):
    c = automation.GetRootControl()
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
    automation.Logger.WriteLine('call python.exe automation_notepad_py3.py', automation.ConsoleColor.DarkGreen)
    while True:
        if None != p.poll():
            break
        if stopEvent.is_set():
            childProcess = []
            for pid, pname in automation.Win32API.EnumProcess():
                ppid = automation.Win32API.GetParentProcessId(pid)
                if ppid == p.pid or pid == p.pid:
                    cmd = automation.Win32API.GetProcessCommandLine(pid)
                    childProcess.append((pid, pname, cmd))
            for pid, pname, cmd in childProcess:
                automation.Logger.WriteLine('kill process: {}, {}, "{}"'.format(pid, pname, cmd), automation.ConsoleColor.Yellow)
                automation.Win32API.TerminateProcess(pid)
            break
        stopEvent.wait(1)
    ReleaseAllKey()
    automation.Logger.WriteLine('test2 exits', automation.ConsoleColor.DarkGreen)

def main():
    automation.RunWithHotKey({
                                   (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_1) : test1
                                   , (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_2) : test2
                               }
                               , (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_4))

if __name__ == '__main__':
    main()
