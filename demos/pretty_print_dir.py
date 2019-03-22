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
        for it in os.scandir(directory):
            absPath = os.path.join(directory, it)
            if os.path.isdir(absPath):
                subdirs.append(absPath)
            else:
                files.append(absPath)
        return subdirs + files


def main(directory, maxDepth = 0xFFFFFFFF):
    remain = {}
    text = []
    absdir = os.path.abspath(directory)
    for it, depth, remainCount in auto.WalkTree(absdir, getChildren=GetDirChildren, includeTop=True, maxDepth=maxDepth):
        remain[depth] = remainCount
        isDir = os.path.isdir(it)
        prefix = ''.join(['│   ' if remain[i+1] else '    ' for i in range(depth - 1)]) # u'│   ' for python 2
        if depth > 0:
            if remain[depth] > 0:
                prefix += '├─→ ' if isDir else '├── '   #'□─→ ' # u'├─→ ' for python 2
            else:
                prefix += '└─→ ' if isDir else '└── '   #'□─→ ' # u'└─→ ' for python 2
        file = it[it.rfind('\\')+1:]
        text.append(prefix)
        text.append(file)
        text.append('\r\n')
        auto.Logger.Write(prefix)
        auto.Logger.WriteLine(file, auto.ConsoleColor.Cyan if os.path.isdir(it) else -1)
    ret = input('\npress Y to save dir tree to clipboard, press other keys to exit')
    if ret.lower() == 'y':
        auto.SetClipboardText(''.join(text))

if __name__ == '__main__':
    if len(sys.argv) == 1:
        dir_ = input('input a dir: ')
        main(dir_)
    elif len(sys.argv) == 2:
        main(sys.argv[1])
    else:
        main(sys.argv[1], int(sys.argv[2]))
