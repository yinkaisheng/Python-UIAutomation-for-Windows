#!python3
# -*- coding: utf-8 -*-
'''
本脚本可以获取QQ2016群所有成员详细资料，请根据提示做对应的操作
作者：yinkaisheng@foxmail.com
2016-01-06
'''
import time
import automation

def GetPersonDetail():
    detailWindow = automation.WindowControl(searchDepth= 1, ClassName = 'TXGuiFoundation', SubName = '的资料')
    detailPane = automation.PaneControl(searchFromControl= detailWindow, Name = '资料')
    details = ''
    for control, depth in automation.WalkTree(detailPane, lambda c: c.GetChildren()):
        if control.ControlType == automation.ControlType.TextControl:
            details += control.Name
        elif control.ControlType == automation.ControlType.EditControl:
            details += control.CurrentValue() + '\n'
    details += '\n' * 2
    detailWindow.Click(0.95, 0.02)
    return details

def main():
    print('请把鼠标放在QQ群聊天窗口中的一个成员上面，3秒后获取\n')
    time.sleep(3)
    listItem = automation.ControlFromCursor()
    if listItem.ControlType != automation.ControlType.ListItemControl:
        print('没有放在群成员上面，程序退出！')
        return
    consoleWindow = automation.GetConsoleWindow()
    consoleWindow.SetActive()
    qqWindow = listItem.GetTopWindow()
    list = listItem.GetParentControl()
    allListItems = list.GetChildren()
    for listItem in allListItems:
        automation.Logger.WriteLine(listItem.Name)
    answer = input('是否获取详细信息？按y和Enter继续\n')
    if answer.lower() == 'y':
        print('\n3秒后开始获取QQ群成员详细资料，您可以一直按住F10键暂停脚本')
        time.sleep(3)
        qqWindow.SetActive()
        time.sleep(0.5)
        #确保群里第一个成员可见在最上面
        left, top, right, bottom = list.BoundingRectangle
        while allListItems[0].BoundingRectangle[1] < top:
            automation.Win32API.MouseClick(right - 5, top + 20)
            time.sleep(0.5)
        for listItem in allListItems:
            if listItem.ControlType == automation.ControlType.ListItemControl:
                if automation.Win32API.IsKeyPressed(automation.Keys.VK_F10):
                    consoleWindow.SetActive()
                    input('\n您暂停了脚本，按Enter继续\n')
                    qqWindow.SetActive()
                    time.sleep(0.5)
                listItem.RightClick()
                menu = automation.MenuControl(searchDepth= 1, ClassName = 'TXGuiFoundation')
                menuItems = menu.GetChildren()
                for menuItem in menuItems:
                    if menuItem.Name == '查看资料':
                        menuItem.Click()
                        break
                automation.Logger.WriteLine(listItem.Name, automation.ConsoleColor.Green)
                automation.Logger.WriteLine(GetPersonDetail())
                listItem.Click()
                time.sleep(0.5)
                automation.SendKeys('{Down}')
                time.sleep(0.5)

if __name__ == '__main__':
    main()
    input('press Enter to exit')
