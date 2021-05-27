#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def main():
    width = 500
    height = 500
    cmdWindow = auto.GetConsoleWindow()
    auto.Logger.WriteLine('create a transparent image')
    bitmap = auto.Bitmap(width, height)
    bitmap.ToFile('image_transparent.png')

    cmdWindow.SetActive()
    auto.Logger.WriteLine('create a blue image')
    start = auto.ProcessTime()
    for x in range(width):
        for y in range(height):
            argb = 0x0000FF | ((255 * x // 500) << 24)
            bitmap.SetPixelColor(x, y, argb)
    cost = auto.ProcessTime() - start
    auto.Logger.WriteLine('write {}x{} image by SetPixelColor cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_blue.png')

    start = auto.ProcessTime()
    for x in range(width):
        for y in range(height):
            bitmap.GetPixelColor(x, y)
    cost = auto.ProcessTime() - start
    auto.Logger.WriteLine('read {}x{} image by GetPixelColor cost {:.3f}s'.format(width, height, cost))

    argb = [(0xFF0000 | (0x0000FF * y // height) | ((255 * x // width) << 24)) for x in range(width) for y in range(height)]
    start = auto.ProcessTime()
    bitmap.SetPixelColorsOfRect(0, 0, width, height, argb)
    cost = auto.ProcessTime() - start
    auto.Logger.WriteLine('write {}x{} image by SetPixelColorsOfRect cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_red.png')

    arrayType = auto.ctypes.c_uint32 * (width * height)
    nativeArray = arrayType(*argb)
    start = auto.ProcessTime()
    bitmap.SetPixelColorsOfRectByNativeArray(0, 0, width, height, nativeArray)
    cost = auto.ProcessTime() - start
    auto.Logger.WriteLine('write {}x{} image by SetPixelColorsOfRectByNativeArray cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_red2.png')

    start = auto.ProcessTime()
    colors = bitmap.GetAllPixelColors()
    cost = auto.ProcessTime() - start
    auto.Logger.WriteLine('read {}x{} image by GetAllPixelColors cost {:.3f}s'.format(width, height, cost))
    bitmap.ToFile('image_red.png')
    bitmap.ToFile('image_red.jpg')

    subprocess.Popen('image_red.png', shell=True)

    root = auto.GetRootControl()
    bitmap = root.ToBitmap(0, 0, 400, 400)
    bitmap.ToFile('desk_part.png')

    width, height = 100, 100
    colors = bitmap.GetPixelColorsOfRects([(0, 0, width, height), (100, 100, width, height), (200, 200, width, height)])
    for i, nativeArray in enumerate(colors):
        bitmap = auto.Bitmap(width, height)
        bitmap.SetPixelColorsOfRectByNativeArray(0, 0, width, height, nativeArray)
        bitmap.ToFile('desk_part{}.png'.format(i))


if __name__ == '__main__':
    main()
