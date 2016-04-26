#!python3
# -*- coding:utf-8 -*-
import os
import time
import subprocess
import automation

def main(stopEvent):
    n = 0
    while True:
        if stopEvent.is_set():
            break
        print(n)
        n += 1
        #time.sleep(1)
        stopEvent.wait(1)
    print('main exits')

def main2(stopEvent):
    p = subprocess.Popen('py.exe automation_notepad_py3.py')
    automation.Logger.WriteLine('call py.exe automation_notepad_py3.py'.format(p.pid), automation.ConsoleColor.DarkGreen)
    while True:
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
    automation.Logger.WriteLine('main exits'.format(p.pid), automation.ConsoleColor.DarkGreen)

if __name__ == '__main__':
    automation.RunWithHotKey(main, (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_1), (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_2))
