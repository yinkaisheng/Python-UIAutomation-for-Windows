#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import threading

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def threadFunc(root):
    """
    If you want to use functionalities related to Controls and Patterns in a new thread.
    You must call InitializeUIAutomationInCurrentThread first in the thread
        and call UninitializeUIAutomationInCurrentThread when the thread exits.
    But you can't use use a Control or a Pattern created in a different thread.
    So you can't create a Control or a Pattern in main thread and then pass it to a new thread and use it.
    """
    #print(root)# you cannot use root because it is root control created in main thread
    th = threading.currentThread()
    auto.Logger.WriteLine('\nThis is running in a new thread. {} {}'.format(th.ident, th.name), auto.ConsoleColor.Cyan)
    time.sleep(2)
    auto.InitializeUIAutomationInCurrentThread()
    auto.GetConsoleWindow().CaptureToImage('console_newthread.png')
    newRoot = auto.GetRootControl()    #ok, root control created in new thread
    auto.EnumAndLogControl(newRoot, 1)
    auto.UninitializeUIAutomationInCurrentThread()
    auto.Logger.WriteLine('\nThread exits. {} {}'.format(th.ident, th.name), auto.ConsoleColor.Cyan)


def main():
    main = threading.currentThread()
    auto.Logger.WriteLine('This is running in main thread. {} {}'.format(main.ident, main.name), auto.ConsoleColor.Cyan)
    time.sleep(2)
    auto.GetConsoleWindow().CaptureToImage('console_mainthread.png')
    root = auto.GetRootControl()
    auto.EnumAndLogControl(root, 1)
    th = threading.Thread(target=threadFunc, args=(root, ))
    th.start()
    th.join()
    auto.Logger.WriteLine('\nMain thread exits. {} {}'.format(main.ident, main.name), auto.ConsoleColor.Cyan)

if __name__ == '__main__':
    main()
