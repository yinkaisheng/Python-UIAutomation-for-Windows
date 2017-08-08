#!python3
# -*- coding: utf-8 -*-
# hide a window with hotkey Ctrl+Shift+1, show the hidden window with hotkey Ctrl+Shift+2
# run notepad.exe first and then press the hotkey for test
import os
import sys
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation

WindowsWantToHide = ('Warcraft III', 'Valve001', 'Counter-Strike', 'Notepad')


def hide():
    root = automation.GetRootControl()
    for window in root.GetChildren():
        if window.ClassName in WindowsWantToHide:
            automation.Logger.WriteLine('hide window, handle {}'.format(window.Handle))
            window.Hide()
            fin = open('hide_windows.txt', 'wt')
            fin.write(str(window.Handle) + '\n')
            fin.close()


def show():
    fout = open('hide_windows.txt')
    lines = fout.readlines()
    fout.close()
    for line in lines:
        handle = int(line)
        window = automation.ControlFromHandle(handle)
        if window:
            automation.Logger.WriteLine('show window: {}'.format(handle))
            window.Show()

def HideWindowFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'python.exe {} hide {}'.format(scriptName, ' '.join(sys.argv[1:]))
    automation.Logger.ColorfulWriteLine('HideWindowFunc call <Color=Green>{}</Color>'.format(cmd))
    p = subprocess.Popen(cmd)
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
    automation.Logger.WriteLine('HideWindowFunc exit')

def ShowWindowFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'python.exe {} show {}'.format(scriptName, ' '.join(sys.argv[1:]))
    automation.Logger.ColorfulWriteLine('ShowWindowFunc call <Color=Green>{}</Color>'.format(cmd))
    p = subprocess.Popen(cmd)
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
    automation.Logger.WriteLine('ShowWindowFunc exit')

if __name__ == '__main__':
    if 'hide' in sys.argv[1:]:
        hide()
    elif 'show' in sys.argv[1:]:
        show()
    else:
        automation.RunWithHotKey({(automation.ModifierKey.MOD_CONTROL|automation.ModifierKey.MOD_SHIFT, automation.Keys.VK_1) : HideWindowFunc
                                    , (automation.ModifierKey.MOD_CONTROL|automation.ModifierKey.MOD_SHIFT, automation.Keys.VK_2) : ShowWindowFunc}
                                 , (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_4)
                                )
