#!python3
# -*- coding:utf-8 -*-
import time
import automation

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
