#!python3
# -*- coding: utf-8 -*-
import os, sys, time

def main():
    os.chdir('dist')
    for file in os.listdir(os.path.curdir):
        if 'py2.' in file:
            newfile = file.replace('py2.', '')
            os.rename(file, newfile)

if __name__ == '__main__':
    main()
