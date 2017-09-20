#!python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation


def main():
    cmdWindow = automation.GetConsoleWindow()
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
    if cmdWindow:
        automation.Logger.ColorfulWriteLine('Scroll to top by SendKeys <Color=Cyan>Ctrl+Home</Color>')
        cmdWindow.SetActive(waitTime = 1)
    mmcWindow.SetActive(waitTime = 1)
    tree.SendKeys('{Ctrl}{Home}', waitTime = 1)
    if cmdWindow:
        automation.Logger.ColorfulWriteLine('Scroll to bottom by SendKeys <Color=Cyan>Ctrl+End</Color>')
        cmdWindow.SetActive(waitTime = 1)
    mmcWindow.SetActive(waitTime = 1)
    tree.SendKeys('{Ctrl}{End}', waitTime = 1)
    if cmdWindow:
        automation.Logger.ColorfulWriteLine('Scroll to top by <Color=Cyan>WheelUp</Color>')
        cmdWindow.SetActive(waitTime = 1)
    print(tree.Handle, tree.Element, len(tree.GetChildren()))
    # before expand, tree has no scrollbar. after expand, tree has a scrollbar.
    # need to Refind on some PCs before find ScrollBarControl from tree
    # maybe the old element has no scrollbar info
    tree.Refind()
    print(tree.Handle, tree.Element, len(tree.GetChildren()))
    vScrollBar = tree.ScrollBarControl(AutomationId = 'NonClientVerticalScrollBar')
    vScrollBarRect = vScrollBar.BoundingRectangle
    thumb = vScrollBar.ThumbControl()
    while True:
        vPercent = tree.CurrentVerticalScrollPercent()
        vPercent2 = vScrollBar.RangeValuePatternCurrentValue()
        print('TreeControl.CurrentVerticalScrollPercent', vPercent)
        print('ScrollBarControl.RangeValuePatternCurrentValue', vPercent2)
        if vPercent2 > 0:
            tree.WheelUp(waitTime = 0.1)
        else:
            break
    if cmdWindow:
        automation.Logger.ColorfulWriteLine('Scroll to bottom by <Color=Cyan>WheelDown</Color>')
        cmdWindow.SetActive(waitTime = 1)
    while True:
        vPercent = tree.CurrentVerticalScrollPercent()
        vPercent2 = vScrollBar.RangeValuePatternCurrentValue()
        print('TreeControl.CurrentVerticalScrollPercent', vPercent)
        print('ScrollBarControl.RangeValuePatternCurrentValue', vPercent2)
        if vPercent2 < 100:
            tree.WheelDown(waitTime = 0.1)
        else:
            break
    if cmdWindow:
        automation.Logger.ColorfulWriteLine('Scroll by <Color=Cyan>DragDrop</Color>')
        cmdWindow.SetActive(waitTime = 1)
    mmcWindow.SetActive(waitTime = 1)
    x, y = thumb.MoveCursorToMyCenter()
    automation.DragDrop(x, y, x, vScrollBarRect[1], waitTime = 1)
    x, y = thumb.MoveCursorToMyCenter()
    automation.DragDrop(x, y, x, vScrollBarRect[3])
    mmcWindow.Close()


if __name__ == '__main__':
    main()
    cmdWindow = automation.GetConsoleWindow()
    if cmdWindow:
        cmdWindow.SetActive()
    input('\npress Enter to exit\n')
