#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import threading

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
from uiautomation import uiautomation as auto


def threadFunc(uselessRoot):
    """
    If you want to use UI Controls in a new thread, create an UIAutomationInitializerInThread object first.
    But you can't use use a Control or a Pattern created in a different thread.
    So you can't create a Control or a Pattern in main thread and then pass it to a new thread and use it.
    """
    # print(uselessRoot)# you cannot use uselessRoot because it is a control created in a different thread
    th = threading.currentThread()
    auto.Logger.WriteLine('\nThis is running in a new thread. {} {}'.format(th.ident, th.name), auto.ConsoleColor.Cyan)
    time.sleep(2)
    with auto.UIAutomationInitializerInThread(debug=True):
        auto.GetConsoleWindow().CaptureToImage('console_newthread.png')
        newRoot = auto.GetRootControl()  # ok, root control created in current thread
        auto.EnumAndLogControl(newRoot, 1)
    auto.Logger.WriteLine('\nThread exits. {} {}'.format(th.ident, th.name), auto.ConsoleColor.Cyan)


def main():
    th = threading.currentThread()
    auto.Logger.WriteLine('This is running in main thread. {} {}'.format(th.ident, th.name), auto.ConsoleColor.Cyan)
    time.sleep(2)
    auto.GetConsoleWindow().CaptureToImage('console_mainthread.png')
    root = auto.GetRootControl()
    auto.EnumAndLogControl(root, 1)
    th = threading.Thread(target=threadFunc, args=(root, ))
    th.start()
    th.join()
    auto.Logger.WriteLine('\nMain thread exits. {} {}'.format(th.ident, th.name), auto.ConsoleColor.Cyan)


if __name__ == '__main__':
    main()
    input('press Enter to exit')