#!python3
# -*- coding:utf-8 -*-
import os
import time

os.environ["PYTHONPATH"] = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # Only required for demo!
from uiautomation import uiautomation as automation


def main():
    consoleWindow = automation.GetConsoleWindow()
    print(consoleWindow)
    print('\nconsole window will be hidden in 3 seconds')
    time.sleep(3)
    consoleWindow.Hide()
    time.sleep(2)
    print('\nconsole window shows again')
    consoleWindow.Show()

if __name__ == '__main__':
    main()
    input('press any key\n')
