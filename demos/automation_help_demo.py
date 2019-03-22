#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess
import ctypes

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def DemoCN():
    """for Chinese language"""
    thisWindow = auto.GetConsoleWindow()
    auto.Logger.ColorfullyWrite('我将运行<Color=Cyan>cmd</Color>并设置它的<Color=Cyan>屏幕缓冲区</Color>使<Color=Cyan>cmd</Color>一行能容纳很多字符\n\n')
    time.sleep(3)

    auto.SendKeys('{Win}r')
    while not isinstance(auto.GetFocusedControl(), auto.EditControl):
        time.sleep(1)
    auto.SendKeys('cmd{Enter}')
    cmdWindow = auto.WindowControl(RegexName = '.+cmd.exe')
    cmdWindow.TitleBarControl().RightClick()
    auto.SendKey(auto.Keys.VK_P)
    optionWindow = cmdWindow.WindowControl(SubName = '属性')
    optionWindow.TabItemControl(SubName = '选项').Click()
    optionTab = optionWindow.PaneControl(SubName = '选项')
    checkBox = optionTab.CheckBoxControl(AutomationId = '103')
    if checkBox.GetTogglePattern().ToggleState != auto.ToggleState.On:
        checkBox.Click()
    checkBox = optionTab.CheckBoxControl(AutomationId = '104')
    if checkBox.GetTogglePattern().ToggleState != auto.ToggleState.On:
        checkBox.Click()
    optionWindow.TabItemControl(SubName = '布局').Click()
    layoutTab = optionWindow.PaneControl(SubName = '布局')
    layoutTab.EditControl(AutomationId='301').GetValuePattern().SetValue('300')
    layoutTab.EditControl(AutomationId='303').GetValuePattern().SetValue('3000')
    layoutTab.EditControl(AutomationId='305').GetValuePattern().SetValue('140')
    layoutTab.EditControl(AutomationId='307').GetValuePattern().SetValue('30')
    optionWindow.ButtonControl(AutomationId = '1').Click()
    cmdWindow.SetActive()
    rect = cmdWindow.BoundingRectangle
    auto.DragDrop(rect.left + 50, rect.top + 10, 50, 30)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('我将运行<Color=Cyan>记事本</Color>并输入<Color=Cyan>Hello!!!</Color>\n\n')
    time.sleep(3)

    subprocess.Popen('notepad')
    notepadWindow = auto.WindowControl(searchDepth = 1, ClassName = 'Notepad')
    cx, cy = auto.GetScreenSize()
    notepadWindow.MoveWindow(cx // 2, 20, cx // 2, cy // 2)
    time.sleep(0.5)
    notepadWindow.EditControl().SendKeys('Hello!!!', 0.05)
    time.sleep(1)

    dir = os.path.dirname(__file__)
    scriptPath = os.path.abspath(os.path.join(dir, '..\\automation.py'))

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('运行"<Color=Cyan>automation.py -h</Color>"显示帮助\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -h'.format(scriptPath) + '{Enter}', 0.05)
    time.sleep(3)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('运行"<Color=Cyan>automation.py -r -d1</Color>"显示所有顶层窗口, 即桌面的子窗口\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -r -d1 -t0'.format(scriptPath) + '{Enter}', 0.05)
    time.sleep(3)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('运行"<Color=Cyan>automation.py -c</Color>"显示鼠标光标下的控件\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -c -t3'.format(scriptPath) + '{Enter}', 0.05)
    notepadWindow.SetActive()
    notepadWindow.MoveCursorToMyCenter()
    time.sleep(3)
    cmdWindow.SetActive(waitTime = 2)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('运行"<Color=Cyan>automation.py -a</Color>"显示鼠标光标下的控件和它的所有父控件\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -a -t3'.format(scriptPath) + '{Enter}', 0.05)
    notepadWindow.SetActive()
    notepadWindow.MoveCursorToMyCenter()
    time.sleep(3)
    cmdWindow.SetActive(waitTime = 2)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('运行"<Color=Cyan>automation.py</Color>"显示当前激活窗口和它的所有子控件\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -t3'.format(scriptPath) + '{Enter}', 0.05)
    notepadWindow.SetActive()
    notepadWindow.EditControl().Click()
    time.sleep(3)
    cmdWindow.SetActive(waitTime = 2)
    time.sleep(3)

    thisWindow.SetActive()
    auto.Logger.WriteLine('演示结束，按Enter退出', auto.ConsoleColor.Green)
    input()


def DemoEN():
    """for other language"""
    thisWindow = auto.GetConsoleWindow()
    auto.Logger.ColorfullyWrite('I will run <Color=Cyan>cmd</Color>\n\n')
    time.sleep(3)

    auto.SendKeys('{Win}r')
    while not isinstance(auto.GetFocusedControl(), auto.EditControl):
        time.sleep(1)
    auto.SendKeys('cmd{Enter}')
    cmdWindow = auto.WindowControl(SubName = 'cmd.exe')
    rect = cmdWindow.BoundingRectangle
    auto.DragDrop(rect.left + 50, rect.top + 10, 50, 10)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('I will run <Color=Cyan>Notepad</Color> and type <Color=Cyan>Hello!!!</Color>\n\n')
    time.sleep(3)

    subprocess.Popen('notepad')
    notepadWindow = auto.WindowControl(searchDepth = 1, ClassName = 'Notepad')
    cx, cy = auto.GetScreenSize()
    notepadWindow.MoveWindow(cx // 2, 20, cx // 2, cy // 2)
    time.sleep(0.5)
    notepadWindow.EditControl().SendKeys('Hello!!!', 0.05)
    time.sleep(1)

    dir = os.path.dirname(__file__)
    scriptPath = os.path.abspath(os.path.join(dir, '..\\automation.py'))

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('run "<Color=Cyan>automation.py -h</Color>" to display the help\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -h'.format(scriptPath) + '{Enter}', 0.05)
    time.sleep(3)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('run "<Color=Cyan>automation.py -r -d1</Color>" to display the top level windows, desktop\'s children\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -r -d1 -t0'.format(scriptPath) + '{Enter}', 0.05)
    time.sleep(3)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('run "<Color=Cyan>automation.py -c</Color>" to display the control under mouse cursor\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -c -t3'.format(scriptPath) + '{Enter}', 0.05)
    notepadWindow.SetActive()
    notepadWindow.MoveCursorToMyCenter()
    time.sleep(3)
    cmdWindow.SetActive(waitTime = 2)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('run "<Color=Cyan>automation.py -a</Color>" to display the control under mouse cursor and its ancestors\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -a -t3'.format(scriptPath) + '{Enter}', 0.05)
    notepadWindow.SetActive()
    notepadWindow.MoveCursorToMyCenter()
    time.sleep(3)
    cmdWindow.SetActive(waitTime = 2)

    thisWindow.SetActive()
    auto.Logger.ColorfullyWrite('run "<Color=Cyan>automation.py</Color>" to display the active window\n\n')
    time.sleep(3)

    cmdWindow.SendKeys('"{}" -t3'.format(scriptPath) + '{Enter}', 0.05)
    notepadWindow.SetActive()
    notepadWindow.EditControl().Click()
    time.sleep(3)
    cmdWindow.SetActive(waitTime = 2)
    time.sleep(3)

    thisWindow.SetActive()
    auto.Logger.WriteLine('press Enter to exit', auto.ConsoleColor.Green)
    input()

if __name__ == '__main__':
    uiLanguage = ctypes.windll.kernel32.GetUserDefaultUILanguage()
    if uiLanguage == 2052:
        DemoCN()
    else:
        DemoEN()
