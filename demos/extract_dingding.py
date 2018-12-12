#!python3
# -*- coding: utf-8 -*-
# author: yinkaisheng@live.com
# extract a company's address book from im.dingtalk.com
import os
import sys
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

SaveFile = os.path.basename(__file__).replace('.py', '.txt')

def StringToInteger(src: str) -> int:
    value, i = 0, 0
    for i, c in enumerate(src):
        if not c.isdigit():
            break
    if i > 0:   value = int(src[:i])
    return value

def CheckPause():
    if auto.Win32API.IsKeyPressed(auto.Keys.VK_F8):
        print('你暂停了脚本，按F9恢复运行')
        while True:
            time.sleep(0.1)
            if auto.Win32API.IsKeyPressed(auto.Keys.VK_F9):
                break

def GetPersonListControl(doc: auto.DocumentControl) -> auto.ListControl:
    return doc.ListControl(searchDepth= 4, Depth = 4, foundIndex= 2)

def GetPersonListAncestorContorl(listControl: auto.ListControl) -> auto.Control:
    return listControl.GetParentControl().GetParentControl()

def ScrollList(listAncestor: auto.Control, personList: auto.Control, direction = 'Bottom'):
    '''direction can be Bottom, Top, Down or Up'''
    prevBottom = personList.BoundingRectangle[3]
    direction = direction.lower()
    rect = listAncestor.BoundingRectangle
    if direction in ['bottom', 'down']:
        WheelFunc = listAncestor.WheelDown
    elif direction in ['top', 'up']:
        WheelFunc = listAncestor.WheelUp
    equalCount = 0
    while True:
        WheelFunc(wheelTimes= 3, waitTime= 0)
        personList.Refind()  # on some Firefox, personList.BoundingRectangle doesn't change after wheel, call Refind to update
        left, top, right, bottom = personList.BoundingRectangle
        if bottom == prevBottom:
            equalCount += 1
            if equalCount == 2:
                break
        else:
            equalCount = 0
        prevBottom = bottom

def ClickListItem(doc: auto.DocumentControl, index: int):
    listControl = GetPersonListControl(doc)
    listAncestor = GetPersonListAncestorContorl(listControl)
    rect = listAncestor.BoundingRectangle
    ScrollList(listAncestor, listControl, 'Top')
    items = listControl.GetChildren()
    item = items[index]
    while item.BoundingRectangle[3] <= 0 or item.BoundingRectangle[3] > rect[3]:
        listAncestor.WheelDown(2)
    item.Click(waitTime = 1)

def GotoParent(doc: auto.DocumentControl) -> int:
    '''return to parent and reutrn the depth of parent'''
    listControl = GetPersonListControl(doc)
    listAncestor = GetPersonListAncestorContorl(listControl)
    ScrollList(listAncestor, listControl, 'Top')
    hierarchyControl = doc.ListControl(searchDepth= 4, Depth = 4, foundIndex= 1)
    items = hierarchyControl.GetChildren()
    if len(items) > 1:
        print('return to parent', items[-2].Name)
        items[-2].Click()
        return len(items) - 2
    return 0

def GetPersonList(doc: auto.DocumentControl) -> list:
    '''
    return a list of tuples (name: str, title: str, info: str)
    if info is not empty, name is a person name, title is person's job title
    else name is a department name, title is member count
    '''
    listControl = GetPersonListControl(doc)
    listAncestor = GetPersonListAncestorContorl(listControl)
    rect = listAncestor.BoundingRectangle
    ScrollList(listAncestor, listControl, 'Bottom')
    time.sleep(1)
    items = listControl.GetChildren()
    ScrollList(listAncestor, listControl, 'Top')
    nodes = []
    departmentCount = 0
    for it in items:
        CheckPause()
        childs = it.GetChildren()
        if len(childs) == 2:
            # deparentment
            nodes.append([childs[0].Name, childs[1].Name, ''])   # department, not a person
        else:
            # person
            while it.BoundingRectangle[3] <= 0 or it.BoundingRectangle[3] > rect[3]:
                listAncestor.WheelDown(2)
                CheckPause()
            names = []
            for c in childs:
                if isinstance(c, auto.EditControl):
                    if c.Name.strip():
                        names.append(c.Name)
            while len(names) < 2:
                names.append('')
            print('click person', names)
            it.Click(waitTime = 1)
            info = ExtractPerson(doc)
            names.append(info)
            nodes.append(names)
    return nodes

def ExtractPerson(doc: auto.DocumentControl) -> list:
    infoList = []
    innerDoc = doc.DocumentControl(searchDepth= 2)
    if innerDoc.Exists(10):
        parent = innerDoc.GetParentControl()
        closeText = parent.TextControl(searchDepth = 3, Name = '')
        lists = innerDoc.GetChildren()
        for item in lists:
            for item2 in item.GetChildren():
                if isinstance(item2, auto.ListControl):
                    childs = item2.GetChildren()
                    if len(childs) >= 2:
                        key = childs[0].Name
                        value = childs[1].Name
                        infoList.append('{}:{}'.format(key, value))
        closeText.Click(simulateMove = False)
    if not infoList:
        auto.Logger.WriteLine('can not get person info', auto.ConsoleColor.Red)
    return infoList

def Parse(doc: auto.DocumentControl, depth = 0):
    nodes = GetPersonList(doc)

    for index, names in enumerate(nodes):
        name, roleOrNum, info = names
        if info:
            # is person
            output = '{}{}[{}][{}]'.format(' ' * (depth * 4), name, roleOrNum, '|'.join(info))
            auto.Logger.WriteLine(output, logFile= SaveFile)

    for index, names in enumerate(nodes):
        name, roleOrNum, info = names
        if not info:
            # is department
            num = StringToInteger(roleOrNum)
            if num > 0:
                output = '{}{}[{}]'.format(' ' * (depth * 4), name, roleOrNum)
                auto.Logger.WriteLine(output, logFile= SaveFile)
                print('click department', names)
                ClickListItem(doc, index)
                Parse(doc, depth + 1)
                GotoParent(doc)

def main():
    input('请先用Firefox登录钉钉，并在组织架构一级目录页面，脚本运行过程中可以按住F8暂停脚本，按F9恢复，现在按Enter键继续')
    consoleWindow = auto.GetConsoleWindow()
    consoleWindow.SetTopmost(True)
    sx, sy = auto.Win32API.GetScreenSize()
    consoleWindow.Move(0, sy // 2)
    consoleWindow.Resize(sx // 3, sy // 2 - 40)
    firefoxWindow = auto.WindowControl(searchDepth = 1, ClassName = 'MozillaWindowClass')
    firefoxWindow.SetActive()
    firefoxWindow.SetWindowVisualState(auto.WindowVisualState.Maximized)
    doc = firefoxWindow.DocumentControl(searchDepth= 4)
    Parse(doc)

if __name__ == '__main__':
    main()
