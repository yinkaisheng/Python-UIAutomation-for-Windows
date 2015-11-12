#!python3
# -*- coding: utf-8 -*-
import os
import automation

def walkDir():
    def GetDirChildren(dir):
        if os.path.isdir(dir):
            subdirs = []
            files = []
            for it in os.listdir(dir):
                absPath = os.path.join(dir, it)
                if os.path.isdir(absPath):
                    subdirs.append(absPath)
                else:
                    files.append(absPath)
            return subdirs + files

    for it, depth in automation.WalkTree(r'c:\Program Files\Internet Explorer', getChildrenFunc= GetDirChildren, includeTop= True):
        print(it, depth)

def walkDesktop():
    def GetFirstChild(control):
        return control.GetFirstChildControl()

    def GetNextSibling(control):
        return control.GetNextSiblingControl()

    desktop = automation.GetRootControl()
    for control, depth in automation.WalkTree(desktop, getFirstChildFunc= GetFirstChild, getNextSiblingFunc= GetNextSibling, includeTop= True, maxDepth= 1):
        print(' ' * depth * 4 + str(control))

def main():
    walkDir()
    print()
    walkDesktop()

if __name__ == '__main__':
    main()
    input('press enter to exit')
