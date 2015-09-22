#!python3
# -*- coding: utf-8 -*-
# rename pdf bookmarks with FoxitReader 6.2.3
import time
import string
import automation

TreeDepth = 2 #书签树只有前两层需要重命名
UpperWords = {
    'amd': 'AMD',
    'api': 'API',
    'apis': 'APIs',
    'arp': 'ARP',
    'dhcp': 'DHCP',
    'dns': 'DNS',
    'e-mail': 'E-mail',
    'e-mails': 'E-mails',
    'eula' : 'EULA',
    'http': 'HTTP',
    'ip': 'IP',
    'mac': 'MAC',
    'unix': 'UNIX',
    'pc': 'PC',
    'pcs': 'PCs',
    'tcp': 'TCP',
    'tcp/ip': 'TCP/IP',
    'vs': 'VS',
    }
LowerWords = ['a', 'an', 'and', 'at', 'for', 'in', 'of', 'on' 'the', 'to', 'with']


class BookMark():
    def __init__(self, name, newName):
        self.name = name
        self.newName = newName
        self.children = []

def main():
    window = automation.WindowControl(searchDepth= 1, ClassName= 'classFoxitReader')
    window.SetActive()
    time.sleep(1)
    tree = automation.TreeControl(searchFromControl= window, ClassName= 'SysTreeView32')
    childItems = tree.GetChildren()
    bookMarks = []
    depth = 1
    for treeItem in childItems:
        if treeItem.ControlType == automation.ControlType.TreeItemControl:
            RenameTreeItem(tree, treeItem, bookMarks, depth)
    fout = open('rename_pdf_bookmark.txt', 'wt', encoding= 'utf-8')
    depth = 1
    for bookMark in bookMarks:
        DumpBookMark(fout, bookMark, depth)
    fout.close()

def DumpBookMark(fout, bookMark, depth):
    fout.write(' ' * (depth - 1) * 4 + bookMark.newName + '\n')
    for child in bookMark.children:
        DumpBookMark(fout, child, depth + 1)

def RenameTreeItem(tree, treeItem, bookMarks, depth):
    treeItem.ScrollIntoView()
    if depth > TreeDepth:
        return
    name = treeItem.Name
    newName = Rename(name)
    bookMark = BookMark(name, newName)
    bookMarks.append(bookMark)
    if newName != name:
        treeItem.RightClick()
        # FoxitReader书签右键菜单(BCGPToolBar，非Windows菜单)弹出后，枚举不到菜单，但从屏幕点上ControlFromPoint能获取到菜单, todo
        # 采用特殊处理获取重命名菜单
        time.sleep(0.2)
        x, y = automation.Win32API.GetCursorPos()
        menuItem = automation.ControlFromPoint(x + 2, y + 2)
        if menuItem.ControlType == automation.ControlType.MenuItemControl:
            #鼠标右下方弹出菜单
            while not (menuItem.Name == '重命名(R)' or menuItem.Name == 'Rename'):
                y += 20
                menuItem = automation.ControlFromPoint(x + 2, y)
        else:
            #鼠标右上方弹出菜单
            menuItem = automation.ControlFromPoint(x + 2, y - 2)
            while not (menuItem.Name == '重命名(R)' or menuItem.Name == 'Rename'):
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
            RenameTreeItem(tree, child, bookMark.children, depth + 1)

def Rename(name):
    newName = name.strip().replace('’', '\'')
    #将CHAPTER 10变成10，删除前置CHAPTER
    if newName.startswith('CHAPTER ') or newName.startswith('Chapter '):
        newName = newName[len('CHAPTER '):]
    newName = newName.title()
    words = newName.split()
    skipIndex = 1 if words[0][-1].isdigit() else 0
    i = 0
    while i < len(words):
        lowerWord = words[i].lower()
        if lowerWord.startswith('www.'):
            words[i] = lowerWord
            i += 1
            continue
        if len(lowerWord) == 1 and lowerWord[0] in string.punctuation:
            if i > 0:
                words[i-1] += lowerWord
                del words[i]
                i -= 1
                continue
        start_punctuation = ''
        end_punctuation = ''
        if lowerWord[0] in string.punctuation:
            start_punctuation = lowerWord[0]
            lowerWord = lowerWord[1:]
        if lowerWord[-1] in string.punctuation:
            end_punctuation = lowerWord[-1]
            lowerWord = lowerWord[:-1]
        if lowerWord in UpperWords:
            words[i] = start_punctuation + UpperWords[lowerWord] + end_punctuation
            i += 1
            continue
        if i > skipIndex and lowerWord in LowerWords:
            if words[i-1][-1] != ':':
                words[i] = lowerWord
        i += 1
    newName = ' '.join(words)
    return newName

if __name__ == '__main__':
    main()
    input('\npress enter to exit')
