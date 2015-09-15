#!python3
# -*- coding: utf-8 -*-
# rename pdf bookmarks with FoxitReader 6.2.3
import time
import string
import automation

TreeDepth = 2 #书签树只有前两层需要重命名
UpperWords = {'tcp' : 'TCP', 'ip': 'IP', 'tcp/ip' : 'TCP/IP', 'dns' : 'DNS', 'pc' : 'PC', 'pcs' : 'PCs'}
LowerWords = ['a', 'an', 'and', 'at', 'in', 'of', 'the', 'to']

def main():
    window = automation.WindowControl(searchDepth= 1, ClassName= 'classFoxitReader')
    window.SetActive()
    time.sleep(1)
    tree = automation.TreeControl(searchFromControl= window, ClassName= 'SysTreeView32')
    childItems = tree.GetChildren()
    depth = 1
    for treeItem in childItems:
        if treeItem.ControlType == automation.ControlType.TreeItemControl:
            RenameTreeItem(tree, treeItem, depth)

def RenameTreeItem(tree, treeItem, depth):
    treeItem.ScrollIntoView()
    if depth > TreeDepth:
        return
    name = treeItem.Name
    newName = name
    #将CHAPTER 10变成10，删除前置CHAPTER
    if newName.startswith('CHAPTER '):
        newName = newName[len('CHAPTER '):]
    newName = newName.title()
    words = newName.split()
    skipIndex = 1 if words[0][-1].isdigit() else 0
    for i in range(len(words)):
        lowerWord = words[i].lower()
        punctuation = ''
        if lowerWord[-1] in string.punctuation:
            punctuation = lowerWord[-1]
            lowerWord = lowerWord[:-1]
        if lowerWord in UpperWords:
            words[i] = UpperWords[lowerWord] + punctuation
            continue
        if i > skipIndex and lowerWord in LowerWords:
            if words[i-1][-1] != ':':
                words[i] = lowerWord
    newName = ' '.join(words)
    if newName != name:
        treeItem.RightClick()
        # FoxitReader书签右键菜单(BCGPToolBar，非Windows菜单)弹出后，枚举不到菜单，但从屏幕点上ControlFromPoint能获取到菜单, todo
        # 采用特殊处理获取重命名菜单
        time.sleep(0.2)
        x, y = automation.Win32API.GetCursorPos()
        menuItem = automation.ControlFromPoint(x + 2, y + 2)
        if menuItem.ControlType == automation.ControlType.MenuItemControl:
            #鼠标右下方弹出菜单
            while menuItem.Name != '重命名(R)':
                y += 20
                menuItem = automation.ControlFromPoint(x + 2, y)
        else:
            #鼠标右上方弹出菜单
            menuItem = automation.ControlFromPoint(x + 2, y - 2)
            while menuItem.Name != '重命名(R)':
                y -= 20
                menuItem = automation.ControlFromPoint(x + 2, y)
        menuItem.Click()
        edit = automation.EditControl(searchFromControl= tree, searchDepth= 1)
        edit.SetValue(newName)
        automation.Win32API.SendKeys('{Enter}')
        print('rename "{0}" to "{1}"'.format(name, newName))
    if depth + 1 > TreeDepth:
        return
    treeItem.Expand()
    childItems = treeItem.GetChildren()
    if childItems:
        treeItem.Expand()
        for child in childItems:
            RenameTreeItem(tree, child, depth + 1)

if __name__ == '__main__':
    main()
    input('press enter to exit')
