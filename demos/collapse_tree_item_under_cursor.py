#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess
import psutil

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def CollapseTreeItem(treeItem: auto.TreeItemControl):
    if not isinstance(treeItem, auto.TreeItemControl):
        return False
    children = treeItem.GetChildren()
    if children:
        for it in children:
            CollapseTreeItem(it)
        ecpt = treeItem.GetExpandCollapsePattern()
        if ecpt and ecpt.ExpandCollapseState == auto.ExpandCollapseState.Expanded:
            ecpt.Collapse(waitTime=0.05)
            return True
    return False

def main():
    treeItem = auto.ControlFromCursor()
    CollapseTreeItem(treeItem)

def HotKeyFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'python.exe {} main {}'.format(scriptName, ' '.join(sys.argv[1:]))
    auto.Logger.WriteLine('call ' + cmd)
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
        stopEvent.wait(0.05)
    auto.Logger.WriteLine('HotKeyFunc exit')


if __name__ == '__main__':
    if 'main' in sys.argv[1:]:
        main()
    else:
        auto.Logger.WriteLine('move mouse to a tree control and press Ctrl+3', auto.ConsoleColor.Green)
        auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_3): HotKeyFunc},
                                 (auto.ModifierKey.Control, auto.Keys.VK_4))

