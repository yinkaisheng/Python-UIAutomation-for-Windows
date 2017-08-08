#!python3
# -*- coding: utf-8 -*-
import os
import sys
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation


def CollapseTreeItem(treeItem):
    if not isinstance(treeItem, automation.TreeItemControl):
        return False
    children = treeItem.GetChildren()
    if children:
        for it in children:
            CollapseTreeItem(it)
        treeItem.Collapse(0)
        return True
    return False


def main():
    treeItem = automation.ControlFromCursor()
    CollapseTreeItem(treeItem)


def HotKeyFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'python.exe {} main {}'.format(scriptName, ' '.join(sys.argv[1:]))
    automation.Logger.WriteLine('call ' + cmd)
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
    automation.Logger.WriteLine('HotKeyFunc exit')


if __name__ == '__main__':
    if 'main' in sys.argv[1:]:
        main()
    else:
        automation.Logger.WriteLine('move mouse to a tree control and press Ctrl+3', automation.ConsoleColor.Green)
        automation.RunWithHotKey({(automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_3) : HotKeyFunc},
                                 (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_4))

