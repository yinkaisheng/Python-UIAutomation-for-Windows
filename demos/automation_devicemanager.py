#!python3
# -*- coding: utf-8 -*-

import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def main():
    sw, sh = auto.GetScreenSize()
    cmdWindow = auto.GetConsoleWindow()
    cmdTransformPattern = cmdWindow.GetTransformPattern()
    cmdTransformPattern.Move(sw // 2, 0)
    cmdTransformPattern.Resize(sw // 2, sh * 3 // 4)
    subprocess.Popen('mmc.exe devmgmt.msc')
    mmcWindow = auto.WindowControl(searchDepth = 1, ClassName = 'MMCMainFrame')
    mmcTransformPattern = mmcWindow.GetTransformPattern()
    mmcTransformPattern.Move(0, 0)
    mmcTransformPattern.Resize(sw // 2, sh * 3 // 4)
    tree = mmcWindow.TreeControl()
    for item, depth in auto.WalkControl(tree, includeTop=True):
        if isinstance(item, auto.TreeItemControl):  #or item.ControlType == auto.ControlType.TreeItemControl
            item.GetSelectionItemPattern().Select(waitTime=0.05)
            pattern = item.GetExpandCollapsePattern()
            if pattern.ExpandCollapseState == auto.ExpandCollapseState.Collapsed:
                pattern.Expand(waitTime=0.05)
            auto.Logger.WriteLine(' ' * (depth - 1) * 4 + item.Name, auto.ConsoleColor.Green)
    if cmdWindow:
        auto.Logger.ColorfullyWriteLine('Scroll by <Color=Cyan>SetScrollPercent</Color>')
        cmdWindow.SetActive(waitTime=1)
    mmcWindow.SetActive(waitTime=1)
    treeScrollPattern = tree.GetScrollPattern()
    treeScrollPattern.SetScrollPercent(auto.ScrollPattern.NoScrollValue, 0)
    treeScrollPattern.SetScrollPercent(auto.ScrollPattern.NoScrollValue, 100)
    if cmdWindow:
        auto.Logger.ColorfullyWriteLine('Scroll to top by SendKeys <Color=Cyan>Ctrl+Home</Color>')
        cmdWindow.SetActive(waitTime=1)
    mmcWindow.SetActive(waitTime = 1)
    tree.SendKeys('{Ctrl}{Home}', waitTime = 1)
    if cmdWindow:
        auto.Logger.ColorfullyWriteLine('Scroll to bottom by SendKeys <Color=Cyan>Ctrl+End</Color>')
        cmdWindow.SetActive(waitTime = 1)
    mmcWindow.SetActive(waitTime = 1)
    tree.SendKeys('{Ctrl}{End}', waitTime = 1)
    if cmdWindow:
        auto.Logger.ColorfullyWriteLine('Scroll to top by <Color=Cyan>WheelUp</Color>')
        cmdWindow.SetActive(waitTime = 1)
    print(tree.NativeWindowHandle, tree.Element, len(tree.GetChildren()))
    # before expand, tree has no scrollbar. after expand, tree has a scrollbar.
    # need to Refind on some PCs before find ScrollBarControl from tree
    # maybe the old element has no scrollbar info
    tree.Refind()
    print(tree.NativeWindowHandle, tree.Element, len(tree.GetChildren()))
    vScrollBar = tree.ScrollBarControl(AutomationId='NonClientVerticalScrollBar')
    rangeValuePattern = vScrollBar.GetRangeValuePattern()
    vScrollBarRect = vScrollBar.BoundingRectangle
    thumb = vScrollBar.ThumbControl()
    while True:
        vPercent = treeScrollPattern.VerticalScrollPercent
        vPercent2 = rangeValuePattern.Value
        print('ScrollPattern.VerticalScrollPercent', vPercent)
        print('ValuePattern.Value', vPercent2)
        if vPercent2 > 0:
            tree.WheelUp(waitTime=0.05)
        else:
            break
    if cmdWindow:
        auto.Logger.ColorfullyWriteLine('Scroll to bottom by <Color=Cyan>WheelDown</Color>')
        cmdWindow.SetActive(waitTime = 1)
    while True:
        vPercent = treeScrollPattern.VerticalScrollPercent
        vPercent2 = rangeValuePattern.Value
        print('ScrollPattern.VerticalScrollPercent', vPercent)
        print('ValuePattern.Value', vPercent2)
        if vPercent2 < 100:
            tree.WheelDown(waitTime = 0.05)
        else:
            break
    if cmdWindow:
        auto.Logger.ColorfullyWriteLine('Scroll by <Color=Cyan>DragDrop</Color>')
        cmdWindow.SetActive(waitTime=1)
    mmcWindow.SetActive(waitTime = 1)
    x, y = thumb.MoveCursorToMyCenter()
    auto.DragDrop(x, y, x, vScrollBarRect.top, waitTime=1)
    x, y = thumb.MoveCursorToMyCenter()
    auto.DragDrop(x, y, x, vScrollBarRect.bottom)
    mmcWindow.GetWindowPattern().Close()


if __name__ == '__main__':
    main()
    cmdWindow = auto.GetConsoleWindow()
    if cmdWindow:
        cmdWindow.SetActive()
    auto.Logger.Write('\nPress any key to exit', auto.ConsoleColor.Cyan)
    import msvcrt
    while not msvcrt.kbhit():
        pass
