#!python3
# -*- coding: utf-8 -*-
"""
本脚本可以获取QQ2018(v9.0)群所有成员详细资料，请根据提示做对应的操作
作者：yinkaisheng@live.com
"""
import os
import sys
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def GetPersonDetail():
    detailWindow = auto.WindowControl(searchDepth= 1, ClassName = 'TXGuiFoundation', SubName = '的资料')
    details = ''
    for control, depth in auto.WalkControl(detailWindow):
        if isinstance(control, auto.EditControl):
            details += control.Name + control.GetValuePattern().Value + '\n'
    details += '\n' * 2
    detailWindow.Click(-10, 10)
    return details


def main():
    auto.Logger.WriteLine('请把鼠标放在QQ群聊天窗口中右下角群成员列表中的一个成员上面，3秒后获取\n', auto.ConsoleColor.Cyan, writeToFile=False)
    time.sleep(3)
    listItem = auto.ControlFromCursor()
    if listItem.ControlType != auto.ControlType.ListItemControl:
        auto.Logger.WriteLine('没有放在群成员上面，程序退出！', auto.ConsoleColor.Cyan, writeToFile=False)
        return
    consoleWindow = auto.GetConsoleWindow()
    if consoleWindow:
        consoleWindow.SetActive()
    qqWindow = listItem.GetTopLevelControl()
    list = listItem.GetParentControl()
    allListItems = list.GetChildren()
    for li in allListItems:
        auto.Logger.WriteLine(li.Name)
        pass
    auto.Logger.WriteLine('是否获取成员详细信息？按F9继续，F10退出', auto.ConsoleColor.Cyan, writeToFile=False)
    while True:
        if auto.IsKeyPressed(auto.Keys.VK_F9):
            break
        elif auto.IsKeyPressed(auto.Keys.VK_F10):
            return
        time.sleep(0.05)
    auto.Logger.WriteLine('\n3秒后开始获取QQ群成员详细资料，您可以一直按住F10键暂停脚本', auto.ConsoleColor.Cyan, writeToFile=False)
    time.sleep(3)
    qqWindow.SetActive()
    #确保群里第一个成员可见在最上面
    list.Click()
    list.SendKeys('{Home}', waitTime = 1)
    for listItem in allListItems:
        if listItem.ControlType == auto.ControlType.ListItemControl:
            if auto.IsKeyPressed(auto.Keys.VK_F10):
                if consoleWindow:
                    consoleWindow.SetActive()
                auto.Logger.WriteLine('\n您暂停了脚本，按F9继续\n', auto.ConsoleColor.Cyan, writeToFile=False)
                while True:
                    if auto.IsKeyPressed(auto.Keys.VK_F9):
                        break
                    time.sleep(0.05)
                qqWindow.SetActive()
            listItem.RightClick(waitTime=2)
            menu = auto.MenuControl(searchDepth= 1, ClassName = 'TXGuiFoundation')
            menuItems = menu.GetChildren()
            for menuItem in menuItems:
                if menuItem.Name == '查看资料':
                    menuItem.Click(40)
                    break
            auto.Logger.WriteLine(listItem.Name, auto.ConsoleColor.Green)
            auto.Logger.WriteLine(GetPersonDetail())
            listItem.Click()
            auto.SendKeys('{Down}')

if __name__ == '__main__':
    main()
    input('press Enter to exit')
