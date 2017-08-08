#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation

PrintTree = False
ExpandFromRoot = False
MaxExpandDepth = 0xFFFFFFFF


def GetTreeItemChildren(item):
    if isinstance(item, automation.TreeItemControl):
        if automation.ExpandCollapseState.Expanded == item.CurrentExpandCollapseState():
            children = item.GetChildren()
            #some tree items need some time to finish expanding
            while not children:
                time.sleep(0.1)
                children = item.GetChildren()
            return children
    else:
        return item.GetChildren()


def ExpandTreeItem(treeItem):
    for item, depth, remainCount in automation.WalkTree(treeItem, getChildrenFunc = GetTreeItemChildren, includeTop = True, maxDepth = MaxExpandDepth):
        if isinstance(item, automation.TreeItemControl):  #or item.ControlType == automation.ControlType.TreeItemControl
            if PrintTree:
                automation.Logger.WriteLine(' ' * (depth * 4) + item.Name)
            if depth < MaxExpandDepth:  # and automation.ExpandCollapseState.Collapsed == item.CurrentExpandCollapseState():
                item.Expand(0)
            item.ScrollIntoView()


def main():
    treeItem = automation.ControlFromCursor()
    if isinstance(treeItem, automation.TreeItemControl) or isinstance(treeItem, automation.TreeControl):
        while ExpandFromRoot:
            if isinstance(treeItem, automation.TreeControl):
                break
            treeItem = treeItem.GetParentControl()
        ExpandTreeItem(treeItem)
    else:
        automation.Logger.WriteLine('the control under cursor is not a tree control', automation.ConsoleColor.Yellow)


def HotKeyFunc(stopEvent):
    cmd = [sys.executable, os.path.abspath(__file__), "--main"] + sys.argv[1:]
    automation.Logger.WriteLine('call {}'.format(cmd))
    p = subprocess.Popen(cmd)
    while True:
        if p.poll() is not None:
            break
        if stopEvent.is_set():
            childProcess = []
            for pid, pname in automation.Win32API.EnumProcess():
                ppid = automation.Win32API.GetParentProcessId(pid)
                if ppid == p.pid or pid == p.pid:
                    cmd = automation.Win32API.GetProcessCommandLine(pid)
                    childProcess.append((pid, pname, cmd))
            for pid, pname, cmd in childProcess:
                automation.Logger.WriteLine('kill process: {}, {}, "{}"'.format(pid, pname, cmd),
                                            automation.ConsoleColor.Yellow)
                automation.Win32API.TerminateProcess(pid)
            break
        stopEvent.wait(1)
    automation.Logger.WriteLine('HotKeyFunc exit')


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--main', action='store_true', help='exec main')
    parser.add_argument('-d', '--depth', type=int, default=0xFFFFFFFF,
                        help='max expand tree depth')
    parser.add_argument('-r', '--root', action='store_true', help='expand from root')
    # print is a keyword in py2, so send to unambiguous "dest".
    parser.add_argument('-p', '--print', action='store_true', help='print tree node text', dest="print_")
    args = parser.parse_args()
    # automation.Logger.WriteLine(str(args))

    if args.main:
        ExpandFromRoot = args.root
        MaxExpandDepth = args.depth
        PrintTree = args.print_
        main()
    else:
        automation.Logger.WriteLine('move mouse to a tree control and press Ctrl+1', automation.ConsoleColor.Green)
        automation.RunWithHotKey({(automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_1): HotKeyFunc},
                                 (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_2))

