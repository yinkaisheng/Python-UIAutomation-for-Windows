#!python3
# -*- coding:utf-8 -*-
import os
import time
import subprocess
import automation

def main(stopEvent):
    n = 0
    while True:
        if stopEvent.is_set():
            break
        print(n)
        n += 1
        #time.sleep(1)
        stopEvent.wait(1)
    print('main exit')

def main2(stopEvent):
    print('call automate_notepad_py3.py')
    #must use 'c:\Python34\python.exe notepadtest3.py', not 'py notepadtest3.py'
    #because py win run c:\Python34\python.exe, these are two processes: py.exe and python.exe
    #p.kill() can only kill py.exe but not python.exe, it can't kill subprocess.
    p = subprocess.Popen(r'c:\Python34\python.exe automate_notepad_py3.py')  # can't use 'py notepadtest3.py'
    print('handle', int(p._handle))
    while True:
        if stopEvent.is_set():
            print('call kill')
            p.kill()
            break
    print('main exit')

if __name__ == '__main__':
    automation.RunWithHotKey(main2, (automation.HotKey.MOD_CONTROL, automation.Keys.VK_1), (automation.HotKey.MOD_CONTROL, automation.Keys.VK_2))
