#!python3
# -*- coding:utf-8 -*-
import sys
import subprocess
from automation import *

if __name__ == '__main__':
    startTime = time.clock()
    isTop = 1
    print(sys.argv)
    if len(sys.argv) == 2:
        isTop = int(sys.argv[1])
    note = WindowControl(searchDepth = 1, ClassName = 'Notepad')
    if note.Exists(0, 0):
        note.SetTopmost(isTop)
    else:
        subprocess.Popen('notepad')
        note.Refind()
        note.SetTopmost(isTop)
    print('done in {0:.3f}s'.format(time.clock()-startTime))
