#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess
import psutil

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

PrintTree = False
ExpandFromRoot = False
MaxExpandDepth = 0xFFFFFFFF


def GetFirstChild(item: auto.Control):
    if isinstance(item, auto.TreeItemControl):
        ecpt = item.GetExpandCollapsePattern()
        if ecpt and ecpt.ExpandCollapseState == auto.ExpandCollapseState.Expanded:
            child = None
            tryCount = 0
            #some tree items need some time to finish expanding
            while not child:
                tryCount += 1
                child = item.GetFirstChildControl()
                if child or tryCount > 20:
                    break
                time.sleep(0.05)
            return child
    else:
        return item.GetFirstChildControl()

def GetNextSibling(item: auto.Control):
    return item.GetNextSiblingControl()

def ExpandTreeItem(treeItem: auto.TreeItemControl):
    for item, depth in auto.WalkTree(treeItem, getFirstChild=GetFirstChild, getNextSibling=GetNextSibling, includeTop=True, maxDepth=MaxExpandDepth):
        if isinstance(item, auto.TreeItemControl):  #or item.ControlType == auto.ControlType.TreeItemControl
            if PrintTree:
                auto.Logger.WriteLine(' ' * (depth * 4) + item.Name, logFile='tree_dump.txt')
            sipt = item.GetScrollItemPattern()
            if sipt:
                sipt.ScrollIntoView(waitTime=0.05)
            ecpt = item.GetExpandCollapsePattern()
            if depth < MaxExpandDepth and ecpt and ecpt.ExpandCollapseState == auto.ExpandCollapseState.Collapsed:  # and auto.ExpandCollapseState.Collapsed == item.CurrentExpandCollapseState():
                ecpt.Expand(waitTime=0.05)


def main():
    tree = auto.ControlFromCursor()
    if ExpandFromRoot:
        tree = tree.GetAncestorControl(lambda c, d: isinstance(c, auto.TreeControl))
    if isinstance(tree, auto.TreeItemControl) or isinstance(tree, auto.TreeControl):
        ExpandTreeItem(tree)
    else:
        auto.Logger.WriteLine('the control under cursor is not a tree control', auto.ConsoleColor.Yellow)


def HotKeyFunc(stopEvent: 'Event', argv: list):
    args = [sys.executable, __file__] + argv
    cmd = ' '.join('"{}"'.format(arg) for arg in args)
    auto.Logger.WriteLine('call {}'.format(cmd))
    p = subprocess.Popen(cmd)
    while True:
        if p.poll() is not None:
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
    parser.add_argument('-d', '--depth', type=int, default=0xFFFFFFFF,
                        help='max expand tree depth')
    parser.add_argument('-r', '--root', action='store_true', help='expand from root')
    parser.add_argument('-p', '--print', action='store_true', help='print tree node text', dest="print_")
    args = parser.parse_args()

    if args.main:
        ExpandFromRoot = args.root
        MaxExpandDepth = args.depth
        PrintTree = args.print_
        main()
    else:
        auto.Logger.WriteLine('move mouse to a tree control and press Ctrl+1', auto.ConsoleColor.Green)
        auto.RunByHotKey({(auto.ModifierKey.Control, auto.Keys.VK_1): lambda event: HotKeyFunc(event, ['--main'] + sys.argv[1:])},
                                 (auto.ModifierKey.Control, auto.Keys.VK_2))

