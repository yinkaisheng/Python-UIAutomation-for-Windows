#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
from uiautomation import uiautomation as auto


def main():

    root = auto.GetRootControl()
    bitmap = root.ToBitmap(0, 0, 400, 400)
    print('save (0,0,400,400) of desktop to desk_part.png')
    bitmap.ToFile('desk_part.png')
    side = int(bitmap.Width * 1.42)
    bitmap2 = auto.Bitmap(side, side)
    bitmap2.Clear(0xFFFF_FFFF)
    bitmap2.Paste(x=(side-bitmap.Width)//2, y=(side-bitmap.Height)//2, bitmap=bitmap)

    num = 20
    bmps = [bitmap2.RotateWithSameSize(bitmap2.Width/2, bitmap2.Height/2, i*360/num) for i in range(0, num)]
    start = auto.ProcessTime()
    auto.TIFF.ToTiffFile('desk_part.tif', bitmaps=bmps)
    cost = auto.ProcessTime() - start
    print('save bitmaps to multiple frames tif cost {:.3f}s'.format(cost))
    start = auto.ProcessTime()
    auto.GIF.ToGifFile('desk_part.gif', bitmaps=bmps, delays=[100]*num)
    cost = auto.ProcessTime() - start
    print('save bitmaps to multiple frames gif cost {:.3f}s'.format(cost))
    # subprocess.Popen('desk_part.gif', shell=True)

    gif: auto.GIF
    gif = auto.Bitmap.FromFile('desk_part.gif')
    print('desk_part.gif is {}'.format(gif))
    for i, bmp in enumerate(gif):
        print('gif frame 0 {}'.format(bmp))
    print('access gif by index 1')
    gif.SelectActiveFrame(1)
    gif.ToFile('gif_index1.jpeg', quality=80)

if __name__ == '__main__':
    main()
