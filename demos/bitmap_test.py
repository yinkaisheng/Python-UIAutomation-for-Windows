#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as automation


def main():
    width = 500
    height = 500
    cmdWindow = automation.GetConsoleWindow()
    automation.Logger.WriteLine('create a transparent image')
    bitmap = automation.Bitmap(width, height)
    bitmap.ToFile('image_transparent.png')

    cmdWindow.SetActive()
    automation.Logger.WriteLine('create a blue image')
    start = time.clock()
    for x in range(width):
        for y in range(height):
            argb = 0x0000FF | ((255 * x // 500) << 24)
            bitmap.SetPixelColor(x, y, argb)
    cost = time.clock() - start
    automation.Logger.WriteLine('write {}x{} image by SetPixelColor cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_blue.png')

    start = time.clock()
    for x in range(width):
        for y in range(height):
            bitmap.GetPixelColor(x, y)
    cost = time.clock() - start
    automation.Logger.WriteLine('read {}x{} image by GetPixelColor cost {:.3f}s'.format(width, height, cost))

    start = time.clock()
    argb = [(0xFF0000 | (0x0000FF * y // height) | ((255 * x // width) << 24)) for x in range(width) for y in range(height)]
    bitmap.SetPixelColorsHorizontally(0, 0, argb)
    cost = time.clock() - start
    automation.Logger.WriteLine('write {}x{} image by SetPixelColorsHorizontally cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_red.png')

    start = time.clock()
    bitmap.GetAllPixelColors()
    cost = time.clock() - start
    automation.Logger.WriteLine('read {}x{} image by GetAllPixelColors cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_red.png')
    bitmap.ToFile('image_red.jpg')

    subprocess.Popen('image_red.png', shell = True)


if __name__ == '__main__':
    main()
