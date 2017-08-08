#!python3
# -*- coding: utf-8 -*-
import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation


def WalkDir():
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

    for it, depth, remainCount in automation.WalkTree(r'c:\Program Files\Internet Explorer', getChildrenFunc=GetDirChildren, includeTop=True):
        print(' ' * depth * 4 + it)


def WalkDesktop():
    def GetFirstChild(control):
        return control.GetFirstChildControl()

    def GetNextSibling(control):
        return control.GetNextSiblingControl()

    desktop = automation.GetRootControl()
    for control, depth in automation.WalkTree(desktop, getFirstChildFunc=GetFirstChild, getNextSiblingFunc=GetNextSibling, includeTop=True, maxDepth= 1):
        print(' ' * depth * 4 + str(control))


def WalkCurrentWindow():
    window = automation.GetForegroundControl().GetTopWindow()
    for control, depth in automation.WalkControl(window, True):
        print(' ' * depth * 4 + str(control))


def main():
    WalkDir()
    print()
    WalkDesktop()
    print()
    WalkCurrentWindow()

if __name__ == '__main__':
    main()
    input('press enter to exit')
