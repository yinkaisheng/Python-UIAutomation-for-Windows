#!python3
# -*- coding: utf-8 -*-
import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

def GetDirChildren(directory):
    if os.path.isdir(directory):
        subdirs = []
        files = []
        for it in os.listdir(directory):
            absPath = os.path.join(directory, it)
            if os.path.isdir(absPath):
                subdirs.append(absPath)
            else:
                files.append(absPath)
        return subdirs + files

def main(directory, maxDepth = 0xFFFFFFFF):
    remain = {}
    texts = []
    absdir = os.path.abspath(directory)
    for it, depth, remainCount in auto.WalkTree(absdir, getChildren=GetDirChildren, includeTop=True, maxDepth=maxDepth):
        remain[depth] = remainCount
        isDir = os.path.isdir(it)
        prefixPrint = ''.join(['│　　' if remain[i + 1] else '　　　' for i in range(depth - 1)])
        prefixLog = ''.join(['│   ' if remain[i + 1] else '    ' for i in range(depth - 1)])
        if depth > 0:
            if remain[depth] > 0:
                prefixPrint += '├─→' if isDir else '├──'
                prefixLog += '├─→ ' if isDir else '├── '
            else:
                prefixPrint += '└─→' if isDir else '└──'
                prefixLog += '└─→ ' if isDir else '└── '
        file = os.path.basename(it)
        texts.append(prefixLog)
        texts.append(file)
        texts.append('\n')
        auto.Logger.Write(prefixPrint, writeToFile=False)
        auto.Logger.WriteLine(file, auto.ConsoleColor.Cyan if isDir else auto.ConsoleColor.Default, writeToFile=False)
    allText = ''.join(texts)
    auto.Logger.WriteLine(allText, printToStdout=False)
    ret = input('\npress Y to save dir tree to clipboard, press other keys to exit')
    if ret.lower() == 'y':
        auto.SetClipboardText(allText)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        dir_ = input('input a dir: ')
        main(dir_)
    elif len(sys.argv) == 2:
        main(sys.argv[1])
    else:
        main(sys.argv[1], int(sys.argv[2]))
