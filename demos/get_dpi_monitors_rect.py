#!python3
# -*- coding: utf-8 -*-
import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def test():
    auto.Logger.WriteLine('\nget screen size dpi unaware:', auto.ConsoleColor.DarkGreen)
    print(auto.GetScreenSize(False))

    auto.Logger.WriteLine('\nget screen size dpi aware per monitor:', auto.ConsoleColor.DarkGreen)
    print(auto.GetScreenSize(True))

    auto.Logger.WriteLine('\nget virtual screen size dpi unaware:', auto.ConsoleColor.DarkGreen)
    print(auto.GetVirtualScreenSize(False))

    auto.Logger.WriteLine('\nget virtual screen size dpi per monitor:', auto.ConsoleColor.DarkGreen)
    print(auto.GetVirtualScreenSize(True))

    auto.Logger.WriteLine('\nget monitors rect dpi unaware:', auto.ConsoleColor.DarkGreen)
    for rect in auto.GetMonitorsRect(False):
        print(rect)

    auto.Logger.WriteLine('\nget monitors rect dpi per monitor:', auto.ConsoleColor.DarkGreen)
    for rect in auto.GetMonitorsRect(True):
        print(rect)

    input('\nPress Enter to exit')

if __name__ == '__main__':
    test()
