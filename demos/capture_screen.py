#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto

def CaptureControl(c, path, up = False):
    if c.CaptureToImage(path):
        auto.Logger.WriteLine('capture image: ' + path)
    else:
        auto.Logger.WriteLine('capture failed', auto.ConsoleColor.Yellow)
    if up:
        r = auto.GetRootControl()
        depth = 0
        name, ext = os.path.splitext(path)
        while True:
            c = c.GetParentControl()
            if not c or auto.ControlsAreSame(c, r):
                break
            depth += 1
            savePath = name + '_p' * depth + ext
            if c.CaptureToImage(savePath):
                auto.Logger.WriteLine('capture image: ' + savePath)
    subprocess.Popen(path, shell = True)


def main(args):
    if args.time > 0:
        time.sleep(args.time)
    start = time.time()
    if args.active:
        c = auto.GetForegroundControl()
        CaptureControl(c, args.path, args.up)
    elif args.cursor:
        c = auto.ControlFromCursor()
        CaptureControl(c, args.path, args.up)
    elif args.fullscreen:
        c = auto.GetRootControl()
        rects = auto.GetMonitorsRect()
        dot = args.path.rfind('.')
        for i, rect in enumerate(rects):
            path = args.path[:dot] + '_' * i + args.path[dot:]
            c.CaptureToImage(path, rect.left, rect.top, rect.width(), rect.height())
            auto.Logger.WriteLine('capture image: ' + path)
    auto.Logger.WriteLine('cost time: {:.3f} s'.format(time.time() - start))


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--fullscreen', action='store_true', default = True,  help = 'capture full screen')
    parser.add_argument('-a', '--active', action='store_true', help = 'capture active window')
    parser.add_argument('-c', '--cursor', action='store_true', help = 'capture control under cursor')
    parser.add_argument('-u', '--up', action='store_true', help = 'capture control under cursor and up to its top window')
    parser.add_argument('-p', '--path', type = str, default = 'capture.png', help = 'save path')
    parser.add_argument('-t', '--time', type = int, default = 0, help = 'delay time')
    args = parser.parse_args()
    auto.Logger.WriteLine(str(args))
    main(args)


