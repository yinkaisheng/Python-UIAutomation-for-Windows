#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import ctypes
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
from uiautomation import uiautomation as auto

def bitmapToPilImageTest(bitmap: auto.Bitmap):
    try:
        from PIL import Image
    except ImportError:
        print('PIL not installed')
        return

    print('bitmap to PIL.Image')
    image = bitmap.ToPILImage()
    print('save png by PIL')
    image.save('pil.png')
    print('save bmp by PIL')
    image.save('pil.bmp')
    print('save jpg by PIL')
    image.convert('RGB').save('pil.jpg', quality=80, optimize=True)

def bitmapToCVImageTest(bitmap: auto.Bitmap):
    try:
        import numpy as np
        import cv2
    except ImportError:
        print('numpy or cv2 not installed')
        return

    print('bitmap to cv2 image')
    bgraImage = bitmap.ToNDArray()
    bgrImage = cv2.cvtColor(bgraImage, cv2.COLOR_BGRA2BGR)
    print('save jpg by cv2')
    cv2.imwrite('cv2.jpg', bgrImage, [cv2.IMWRITE_JPEG_QUALITY, 80])

def main():
    root = auto.GetRootControl()
    bitmap = root.ToBitmap(0, 0, 400, 400)
    print('save {} (0,0,400,400) of desktop to desk_part.png'.format(bitmap))
    bitmap.ToFile('desk_part.png')
    bitmapToPilImageTest(bitmap)
    bitmapToCVImageTest(bitmap)

if __name__ == '__main__':
    main()
