#!python3
# -*- coding:utf-8 -*-
import os
import automation

def main():
    print('main')
    os.system('py -3 rename_pdf_bookmark.py')

if __name__ == '__main__':
    automation.RunHotKey(main, (automation.HotKey.MOD_ALT, automation.Keys.VK_1), (automation.HotKey.MOD_ALT, automation.Keys.VK_2))
