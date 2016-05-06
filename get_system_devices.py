#!python3
# -*- coding: utf-8 -*-

import time
import subprocess
import automation

def main():
    subprocess.Popen('mmc.exe devmgmt.msc')
    mmcWindow = automation.WindowControl(searchDepth = 1, ClassName = 'MMCMainFrame')
    tree = automation.TreeControl(searchFromControl = mmcWindow)
    for item, depth in automation.WalkTree(tree, getChildrenFunc = lambda c : c.GetChildren(), includeTop = True):
        if isinstance(item, automation.TreeItemControl):  #or item.ControlType == automation.ControlType.TreeItemControl
            item.Select()
            if automation.ExpandCollapseState.Collapsed == item.CurrentExpandCollapseState():
                item.Expand()
            automation.Logger.WriteLine(' ' * (depth - 1) * 4 + item.Name, automation.ConsoleColor.Green)
            time.sleep(0.1)

if __name__ == '__main__':
    main()
    automation.GetConsoleWindow().SetActive()
    input('\npress Enter to exit\n')
