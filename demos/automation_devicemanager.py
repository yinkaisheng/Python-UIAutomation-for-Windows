#!python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def main():
    sw, sh = auto.Win32API.GetScreenSize()
    cmdWindow = auto.GetConsoleWindow()
    subprocess.Popen('mmc.exe devmgmt.msc')
    time.sleep(1)
    mmcWindow = auto.WindowControl(searchDepth = 1, ClassName = 'MMCMainFrame')
    mmcWindow.Move(0, 0)
    mmcWindow.Resize(sw // 2, sh * 3 // 4)
    cmdWindow.Move(sw // 2, 0)
    cmdWindow.Resize(sw // 2, sh * 3 // 4)
    tree = mmcWindow.TreeControl()
    for item, depth in auto.WalkControl(tree, includeTop=True):
        if isinstance(item, auto.TreeItemControl):  #or item.ControlType == auto.ControlType.TreeItemControl
            item.Select()
            if auto.ExpandCollapseState.Collapsed == item.CurrentExpandCollapseState():
                item.Expand(0)
            auto.Logger.WriteLine(' ' * (depth - 1) * 4 + item.Name, auto.ConsoleColor.Green)
            time.sleep(0.05)
    if cmdWindow:
        auto.Logger.ColorfulWriteLine('Scroll by <Color=Cyan>SetScrollPercent</Color>')
        cmdWindow.SetActive(waitTime=1)
    mmcWindow.SetActive(waitTime=1)
    for v in range(100, -1, -10):
        tree.SetScrollPercent(0, v)
        time.sleep(0.05)
    for v in range(0, 101, 10):
        tree.SetScrollPercent(0, v)
        time.sleep(0.05)
    if cmdWindow:
        auto.Logger.ColorfulWriteLine('Scroll to top by SendKeys <Color=Cyan>Ctrl+Home</Color>')
        cmdWindow.SetActive(waitTime = 1)
    mmcWindow.SetActive(waitTime = 1)
    tree.SendKeys('{Ctrl}{Home}', waitTime = 1)
    if cmdWindow:
        auto.Logger.ColorfulWriteLine('Scroll to bottom by SendKeys <Color=Cyan>Ctrl+End</Color>')
        cmdWindow.SetActive(waitTime = 1)
    mmcWindow.SetActive(waitTime = 1)
    tree.SendKeys('{Ctrl}{End}', waitTime = 1)
    if cmdWindow:
        auto.Logger.ColorfulWriteLine('Scroll to top by <Color=Cyan>WheelUp</Color>')
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
            tree.WheelUp(waitTime=0.05)
        else:
            break
    if cmdWindow:
        auto.Logger.ColorfulWriteLine('Scroll to bottom by <Color=Cyan>WheelDown</Color>')
        cmdWindow.SetActive(waitTime = 1)
    while True:
        vPercent = tree.CurrentVerticalScrollPercent()
        vPercent2 = vScrollBar.RangeValuePatternCurrentValue()
        print('TreeControl.CurrentVerticalScrollPercent', vPercent)
        print('ScrollBarControl.RangeValuePatternCurrentValue', vPercent2)
        if vPercent2 < 100:
            tree.WheelDown(waitTime = 0.05)
        else:
            break
    if cmdWindow:
        auto.Logger.ColorfulWriteLine('Scroll by <Color=Cyan>DragDrop</Color>')
        cmdWindow.SetActive(waitTime=1)
    mmcWindow.SetActive(waitTime = 1)
    x, y = thumb.MoveCursorToMyCenter()
    auto.DragDrop(x, y, x, vScrollBarRect[1], waitTime = 1)
    x, y = thumb.MoveCursorToMyCenter()
    auto.DragDrop(x, y, x, vScrollBarRect[3])
    mmcWindow.Close()


if __name__ == '__main__':
    main()
    cmdWindow = auto.GetConsoleWindow()
    if cmdWindow:
        cmdWindow.SetActive()
    auto.Logger.WriteLine('\npress Enter to exit\n')
