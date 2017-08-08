#!python3
# -*- coding: utf-8 -*-
#rename pdf bookmarks with FoxitReader 7.2.4, 参考: http://www.cnblogs.com/Yinkaisheng/p/4820954.html
import os
import sys
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation

TreeDepth = 1  #书签树需要命名的层数
TreeSkipDepthList = []
RenameFunction = None
LowerWords = ['a', 'an', 'and', 'at', 'for', 'in', 'of', 'on', 'the', 'to', 'with', 'from']
MatchWords = ['chapter', 'ch', 'section', 'hour', 'lession', 'tutorial']#, 'step'
Punctuation = '!"\',:;?)'  # '!"\',.:;?)'
PunctuationPostfix = '!,:;?)'
Renamed = False
UpperWords = {
    'amd': 'AMD',
    'api': 'API',
    'apis': 'APIs',
    'arp': 'ARP',
    'dhcp': 'DHCP',
    'dns': 'DNS',
    'e-mail': 'E-mail',
    'e-mails': 'E-mails',
    'ee': 'EE',
    'eula' : 'EULA',
    'http': 'HTTP',
    'ip': 'IP',
    'mac': 'MAC',
    'unix': 'UNIX',
    'pc': 'PC',
    'pcs': 'PCs',
    'tcp': 'TCP',
    'tcp/ip': 'TCP/IP',
    'udp': 'UDP',
    'uml': 'UML',
    'i': 'I',
    'ii': 'II',
    'iii': 'III',
    'iv': 'IV',
    'v': 'V',
    'vi': 'VI',
    'vii': 'VII',
    'viii': 'VIII',
    'ix': 'IX',
    'x': 'X',
    'xi': 'XI',
    'xii': 'XII',
    'xiii': 'XIII',
    'xiv': 'XIV',
    'xv': 'XV',
    }


class BookMark():
    def __init__(self, name, newName):
        self.name = name
        self.newName = newName
        self.children = []


def BatchRename():
    foxitWindow = automation.WindowControl(searchDepth= 1, ClassName= 'classFoxitReader')
    foxitWindow.ShowWindow(automation.ShowWindow.ShowMaximized)
    ret = foxitWindow.SetActive()
    print('SetActive', ret)
    automation.Logger.Log(foxitWindow.Name[:-len(' - Foxit Reader')] + '\n')
    automation.SendKeys('{Ctrl}0')
    for aid in ['59583', '60682', '60683']:
        editToolBar = automation.ToolBarControl(searchFromControl= foxitWindow, AutomationId = aid, Name = 'Caption Bar')
        if editToolBar.Exists(0, 0):
            editToolBar.Click(-40, 0.5)
            automation.SendKeys('{Alt}y')
    paneWindow = automation.WindowControl(searchFromControl= foxitWindow, AutomationId = '65280')
    bookmarkPane = automation.PaneControl(searchFromControl= paneWindow, searchDepth= 1, foundIndex= 1)
    l, t, r, b = bookmarkPane.BoundingRectangle
    #bookmarkButton = automation.ButtonControl(searchFromControl= bookmarkPane, Name = '书签') # can't find, but automation -a can find it, why
    if bookmarkPane.Name == '书签':
        if r - l < 40:
            bookmarkButton = automation.ControlFromPoint(l + 10, t + 40)
            if bookmarkButton.Name == '书签':
                bookmarkButton.Click(simulateMove= False)
    else:
        bookmarkButton = automation.ControlFromPoint(l + 10, t + 40)
        if bookmarkButton.Name == '书签':
            bookmarkButton.Click(simulateMove= False)
    tree = automation.TreeControl(searchFromControl= foxitWindow, ClassName= 'SysTreeView32')
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
    if Renamed:
        automation.Logger.Log('rename pdf: ' + foxitWindow.Name)
        automation.SendKeys('{Ctrl}s')
        while '*' in foxitWindow.Name:
            time.sleep(0.5)


def DumpBookMark(fout, bookMark, depth):
    fout.write(' ' * (depth - 1) * 4 + bookMark.newName + '\n')
    for child in bookMark.children:
        DumpBookMark(fout, child, depth + 1)


def RenameTreeItem(tree, treeItem, bookMarks, depth, removeChapter = True):
    treeItem.ScrollIntoView()
    if depth > TreeDepth:
        return
    name = treeItem.Name
    if not name.strip():
        return
    newName, alert = RenameFunction(name, removeChapter)
    if newName.startswith('Appendix'):
        removeChapter = False
    bookMark = BookMark(name, newName)
    bookMarks.append(bookMark)
    if newName != name and depth not in TreeSkipDepthList:
        global Renamed
        Renamed = True
        time.sleep(0.1)
        find = False
        for i in range(3):
            treeItem.RightClick(simulateMove = False, waitTime = 0.2)
            # FoxitReader书签右键菜单(BCGPToolBar，非Windows菜单)弹出后，枚举不到菜单，但从屏幕点上ControlFromPoint能获取到菜单, todo
            # 采用特殊处理获取重命名菜单
            x, y = automation.Win32API.GetCursorPos()
            menuItem = automation.ControlFromPoint(x + 2, y + 2)
            if isinstance(menuItem, automation.MenuItemControl):
                #鼠标右下方弹出菜单
                while 1:
                    if menuItem.Name == '重命名(R)' or menuItem.Name == 'Rename':
                        find = True
                        break
                    y += 20
                    menuItem = automation.ControlFromPoint(x + 2, y)
                    if not isinstance(menuItem, automation.MenuItemControl):
                        break
            else:
                #鼠标右上方弹出菜单
                menuItem = automation.ControlFromPoint(x + 2, y - 2)
                while 1:
                    if menuItem.Name == '重命名(R)' or menuItem.Name == 'Rename':
                        find = True
                        break
                    y -= 20
                    menuItem = automation.ControlFromPoint(x + 2, y)
                    if not isinstance(menuItem, automation.MenuItemControl):
                        break
            if find:
                break
            else:
                automation.SendKeys('{Esc}')
                time.sleep(0.5)
        if not find:
            automation.Logger.Log('this pdf not support editing')
            exit(0)
        menuItem.Click(simulateMove = False, waitTime = 0.2)
        edit = automation.EditControl(searchFromControl= tree, searchDepth= 1)
        edit.SetValue(newName, waitTime = 0.05)
        automation.SendKeys('{Enter}', waitTime = 0.05)
        automation.Logger.Write('rename: ')
        automation.Logger.WriteLine(name, automation.ConsoleColor.Green)
        automation.Logger.Write('    @@to : ' if alert else '    to: ')
        automation.Logger.WriteLine(newName, automation.ConsoleColor.Yellow if alert else automation.ConsoleColor.Green)
    if depth + 1 > TreeDepth:
        return
    if automation.ExpandCollapseState.Collapsed == treeItem.CurrentExpandCollapseState():
        treeItem.Expand(0.1)
    childItems = treeItem.GetChildren()
    if childItems:
        for child in childItems:
            RenameTreeItem(tree, child, bookMark.children, depth + 1, removeChapter)


def Rename1(name, removeChapter = True):
    alert = False
    newName = name.strip().replace('‘', '\'').replace('’', '\'').replace('“', '"').replace('”', '"')
    #newName = newName.replace('.', ' ')
    #newName = newName.replace('_ _', '__')
    #newName = newName.replace('_  _', '__')
    words = newName.split()
    if len(words) == 0:
        return '', alert
    firstWord = words[0].lower()
    if removeChapter and firstWord in MatchWords and len(words) > 2:
        del words[0]
    i = 0
    while i < len(words):
        if words[i][0] in Punctuation:
            if i > 0 and len(words[i]) == 1:
                alert = True
                words[i-1] += words[i]
                del words[i]
                i -= 1
                continue
        i += 1
    newName = ' '.join(words)
    return newName, alert


def Rename2(name, removeChapter = True):
    alert = False
    newName = name.strip().replace('‘', '\'').replace('’', '\'').replace('“', '"').replace('”', '"')
    if newName.startswith('www.'):
        return newName, alert
    newName = newName.title()
    words = newName.split()
    if len(words) == 0:
        return '', alert
    firstWord = words[0].lower()
    if removeChapter and firstWord in MatchWords and len(words) > 2:
        del words[0]
    skipIndex = 1 if words[0][-1].isdigit() else 0
    i = 0
    while i < len(words):
        lowerWord = words[i].lower()
        if lowerWord[0] in PunctuationPostfix:
            if i > 0 and len(words[i]) == 1:
                alert = True
                words[i-1] += words[i]
                del words[i]
                i -= 1
                continue
        start_punctuation = ''
        end_punctuation = ''
        if lowerWord[0] in Punctuation:
            start_punctuation = lowerWord[0]
            lowerWord = lowerWord[1:]
        if lowerWord[-1] in Punctuation:
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
    return newName, alert

def usage():
    sys.stdout.write("""usage
-h      show command help
-r      rename function, 1 or 2
-d      bookmark tree depth
""")


def main():
    BatchRename()

if __name__ == '__main__':
    import getopt
    sys.stdout.write(str(sys.argv) + '\n')
    options, args = getopt.getopt(sys.argv[1:], 'hr:d:s:',
                                  ['help', 'rename=', 'depth=', 'skip='])
    RenameFunction = Rename1
    for (o, v) in options:
        if o in ('-h', '-help'):
            usage()
            exit(0)
        elif o in ('-r', '-rename'):
            rename = int(v)
            if rename == 2:
                RenameFunction = Rename2
        elif o in ('-d', '-depth'):
            TreeDepth = int(v)
        elif o in ('-s', '-skip'):
            skipDepth = v.split(', ')
            TreeSkipDepthList = [int(it) for it in skipDepth]
    main()
    #input('\npress enter to exit')
