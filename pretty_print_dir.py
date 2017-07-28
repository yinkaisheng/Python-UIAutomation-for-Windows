#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import uiautomation as automation

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

def main(directory):
    abs_dir = os.path.abspath(directory)
    for it, depth in automation.WalkTree(abs_dir, getChildrenFunc= GetDirChildren, includeTop= True):
        if depth == 0:
            automation.Logger.WriteLine(it[it.rfind('\\')+1:], automation.ConsoleColor.Cyan)
        else:
            if os.path.isdir(it):
                automation.Logger.ColorfulWriteLine('|     ' * (depth - 1) + '|---- <Color=Cyan>' + it[it.rfind('\\')+1:] + '</Color>')
            else:
                automation.Logger.WriteLine('|     ' * (depth - 1) + '|---- ' + it[it.rfind('\\')+1:])

if __name__ == '__main__':
    if len(sys.argv) == 1:
        adir = input('input a dir: ')
        if not adir:
            adir = '.'
        main(adir)
    else:
        main(sys.argv[1])
    input('press Enter to exit')
