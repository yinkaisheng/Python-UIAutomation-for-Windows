#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess
import automation

PrintTree = False
ExpandFromRoot = False
MaxExpandDepth = 0xFFFFFFFF

def getKeyName(theDict, theValue):
    for key in theDict:
        if theValue == theDict[key]:
            return key

def ExpandTreeItemRecursively(treeItem, depth = 0):
    if isinstance(treeItem, automation.TreeItemControl):
        if PrintTree:
            automation.Logger.WriteLine(' ' * (depth * 4) + treeItem.Name)
        if depth >= MaxExpandDepth:
            return
        children = []
        state = treeItem.CurrentExpandCollapseState()
        #print(treeItem.Name, getKeyName(automation.ExpandCollapseState.__dict__, state))
        if state == automation.ExpandCollapseState.Collapsed:
            treeItem.Expand()
        elif state == automation.ExpandCollapseState.LeafNode:
            return
        children = treeItem.GetChildren()
        #some tree items need some time to finish expanding
        tryCount = 0
        while not children:
            if tryCount > 10:
                break
            time.sleep(0.1)
            children = treeItem.GetChildren()
            tryCount += 1
        for it in children:
            it.ScrollIntoView()
            ExpandTreeItemRecursively(it, depth + 1)

def ExpandTreeItem(treeItem):
    for item, depth in automation.WalkTree(treeItem, getChildrenFunc = lambda c : c.GetChildren(), includeTop = True, maxDepth = MaxExpandDepth):
        if isinstance(item, automation.TreeItemControl):  #or item.ControlType == automation.ControlType.TreeItemControl
            #item.Select()
            if depth < MaxExpandDepth and automation.ExpandCollapseState.Collapsed == item.CurrentExpandCollapseState():
                item.Expand()
            item.ScrollIntoView()
            if PrintTree:
                automation.Logger.WriteLine(' ' * (depth * 4) + item.Name)

def main():
    treeItem = automation.ControlFromCursor()
    if isinstance(treeItem, automation.TreeItemControl) or isinstance(treeItem, automation.TreeControl):
        while ExpandFromRoot:
            pc = treeItem.GetParentControl()
            if isinstance(pc, automation.TreeControl) or not pc:
                break
            treeItem = pc
        ExpandTreeItemRecursively(treeItem)
    else:
        automation.Logger.WriteLine('the control under cursor is not a tree control', automation.ConsoleColor.Yellow)


def HotKeyFunc(stopEvent):
    scriptName = os.path.basename(__file__)
    cmd = r'py.exe {} main {}'.format(scriptName, ' '.join(sys.argv[1:]))
    print('call ' + cmd)
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
    print('HotKeyFunc exit')


if __name__ == '__main__':
    if 'main' in sys.argv[1:]:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument('main', help = 'exec main')
        parser.add_argument('-d', '--depth', type = int, default = 0xFFFFFFFF,
                            help = 'max expand tree depth')
        parser.add_argument('-r', '--root', action='store_true', help = 'expand from root')
        parser.add_argument('-p', '--print', action='store_true', help = 'print tree node text')
        args = parser.parse_args()
        automation.Logger.WriteLine(str(args))
        ExpandFromRoot = args.root
        MaxExpandDepth = args.depth
        PrintTree = args.print
        main()
    else:
        automation.Logger.WriteLine('move mouse to a tree control and press Ctrl+1', automation.ConsoleColor.Green)
        automation.RunWithHotKey(HotKeyFunc, (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_1), (automation.ModifierKey.MOD_CONTROL, automation.Keys.VK_2))

