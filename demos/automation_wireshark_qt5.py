#!python3
# -*- coding: utf-8 -*-
"""
Tested for Wireshark 3.0.
Just for fun.
Wireshark can export packets to csv or txt. This is the fastest way.
"""
import os
import sys
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def walk():
    wiresharkWindow = None
    for win in auto.GetRootControl().GetChildren():
        if win.ClassName == 'MainWindow' and win.AutomationId == 'MainWindow':
            if win.ToolBarControl(AutomationId='MainWindow.displayFilterToolBar').Exists(0, 0):
                wiresharkWindow = win
                break
    if not wiresharkWindow:
        auto.Logger.WriteLine('Can not find Wireshark', auto.ConsoleColor.Yellow)
        return

    console = auto.GetConsoleWindow()
    if console:
        sx, sy = auto.GetScreenSize()
        tp = console.GetTransformPattern()
        tp.Resize(sx, sy // 4)
        tp.Move(0, sy - sy // 4 - 50)
        console.SetTopmost()

    wiresharkWindow.SetActive(waitTime=0.1)
    wiresharkWindow.Maximize()
    auto.Logger.ColorfullyWriteLine('Press <Color=Cyan>F1</Color> to stop', auto.ConsoleColor.Yellow)
    tree = wiresharkWindow.TreeControl(searchDepth=4, ClassName='PacketList', Name='Packet list')
    rect = tree.BoundingRectangle
    tree.Click(y=50, waitTime=0.1)
    auto.SendKeys('{Home}', waitTime=0.1)
    columnCount = 0
    treeItemCount = 0
    for item, depth in auto.WalkControl(tree):
        if isinstance(item, auto.HeaderControl):
            columnCount += 1
            auto.Logger.Write(item.Name + ' ')
        elif isinstance(item, auto.TreeItemControl):
            if treeItemCount % columnCount == 0:
                auto.Logger.Write('\n')
                time.sleep(0.1)
            treeItemCount += 1
            auto.Logger.Write(item.Name + ' ')
            if item.BoundingRectangle.bottom >= rect.bottom:
                auto.SendKeys('{PageDown}', waitTime=0.1)
        if auto.IsKeyPressed(auto.Keys.VK_F1):
            auto.Logger.WriteLine('\nF1 pressed', auto.ConsoleColor.Yellow)
            break


if __name__ == '__main__':
    walk()
    time.sleep(2)
