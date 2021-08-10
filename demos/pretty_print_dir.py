#!python3
# -*- coding: utf-8 -*-
import os
import sys
from typing import Tuple
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def listDir(path: Tuple[str, bool]):
    if path[1]:
        files = []
        files2 = []
        for it in os.listdir(path[0]):
            absPath = os.path.join(path[0], it)
            if os.path.isdir(absPath):
                files.append((absPath, True))
            else:
                files2.append((absPath, False))
        files.extend(files2)
        return files


def main(directory: str, maxDepth: int = 0xFFFFFFFF):
    remain = {}
    texts = []
    absdir = os.path.abspath(directory)
    for (it, isDir), depth, remainCount in auto.WalkTree((absdir, True), getChildren=listDir, includeTop=True, maxDepth=maxDepth):
        remain[depth] = remainCount
        prefix = ''.join('┃   ' if remain[i] else '    ' for i in range(1, depth))
        if depth > 0:
            if remain[depth] > 0:
                prefix += '┣━> ' if isDir else '┣━━ '
            else:
                prefix += '┗━> ' if isDir else '┗━━ '
        file = os.path.basename(it)
        texts.append(prefix)
        texts.append(file)
        texts.append('\n')
        auto.Logger.Write(prefix, writeToFile=False)
        auto.Logger.WriteLine(file, auto.ConsoleColor.Cyan if isDir else auto.ConsoleColor.Default, writeToFile=False)
    allText = ''.join(texts)
    auto.Logger.WriteLine(allText, printToStdout=False)
    ret = input('\npress Y to save dir tree to clipboard, press other keys to exit\n')
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
