#!python3
# -*- coding: utf-8 -*-

import os
import time
import subprocess

os.environ["PYTHONPATH"] = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Only required for demo!
from uiautomation import uiautomation as automation


def main():
    subprocess.Popen('mmc.exe devmgmt.msc')
    time.sleep(1)
    mmcWindow = automation.WindowControl(searchDepth = 1, ClassName = 'MMCMainFrame')
    tree = mmcWindow.TreeControl()
    for item, depth in automation.WalkControl(tree, includeTop = True):
        if isinstance(item, automation.TreeItemControl):  #or item.ControlType == automation.ControlType.TreeItemControl
            item.Select()
            if automation.ExpandCollapseState.Collapsed == item.CurrentExpandCollapseState():
                item.Expand(0)
            automation.Logger.WriteLine(' ' * (depth - 1) * 4 + item.Name, automation.ConsoleColor.Green)
            time.sleep(0.1)

if __name__ == '__main__':
    main()
    automation.GetConsoleWindow().SetActive()
    input('\npress Enter to exit\n')
