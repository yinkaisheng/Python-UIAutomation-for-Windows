#!python3
# -*- coding:utf-8 -*-
import os
import sys
import subprocess
import psutil

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def test1(stopEvent):
    """
    This function runs in a new thread triggered by hotkey.
    """
    auto.InitializeUIAutomationInCurrentThread()
    n = 0
    child = None
    auto.Logger.WriteLine('Use UIAutomation in another thread:', auto.ConsoleColor.Yellow)
    while True:
        if stopEvent.is_set():
            break
        if not child:
            n = 1
            child = auto.GetRootControl().GetFirstChildControl()
        auto.Logger.WriteLine(n, auto.ConsoleColor.Cyan)
        auto.LogControl(child)
        child = child.GetNextSiblingControl()
        n += 1
        stopEvent.wait(1)
    auto.UninitializeUIAutomationInCurrentThread()
    print('test1 exits')


def test2(stopEvent):
    """This function runs in a thread triggered by hotkey"""
    cmd = '"{}" "automation_notepad.py"'.format(sys.executable)
    p = subprocess.Popen(cmd)
    auto.Logger.WriteLine(cmd, auto.ConsoleColor.DarkGreen)
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
    auto.Logger.Write('\nUse UIAutomation in main thread:\n', auto.ConsoleColor.Yellow)
    auto.LogControl(auto.GetRootControl())
    auto.Logger.Write('\n')
    auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): test1,
                              (auto.ModifierKey.Control, auto.Keys.VK_2): test2},
                             (auto.ModifierKey.Control, auto.Keys.VK_4))
    auto.SetConsoleTitle(oriTitle)

if __name__ == '__main__':
    main()
