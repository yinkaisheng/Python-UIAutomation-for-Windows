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
        stopEvent.wait(0.05)
    auto.Logger.WriteLine('HotKeyFunc exit')


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--main', action='store_true', help='exec main')
    args = parser.parse_args()

    if args.main:
        main()
    else:
        auto.Logger.WriteLine('move mouse to a tree control and press Ctrl+3', auto.ConsoleColor.Green)
        auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_3): lambda event: HotKeyFunc(event, ['--main'])},
                                 (auto.ModifierKey.Control, auto.Keys.VK_4))

