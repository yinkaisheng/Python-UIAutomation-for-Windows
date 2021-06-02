#!python3
# -*- coding: utf-8 -*-
import os
import sys
import time
import random
from threading import Thread
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
import uiautomation as auto


def threadFunc(num: int):
    count = 0
    end = auto.ProcessTime() + 2
    while auto.ProcessTime() < end:
        count += 1
        if not auto.SetClipboardText('{}'.format(random.randint(0, 10000))):
            print(num, 'failed')
    print(num, 'count', count)


def testThread():
    th1 = Thread(None, threadFunc, args=(1, ))
    th1.start()
    th2 = Thread(None, threadFunc, args=(2, ))
    th2.start()
    time.sleep(2)
    th1.join()
    th2.join()


def main():
    formats = auto.GetClipboardFormats()
    for k, v in formats.items():
        auto.Logger.WriteLineColorfully('Clipboard has format <Color=Cyan>{}, {}</Color>'.format(k, v))
        if k == auto.ClipboardFormat.CF_UNICODETEXT:
            auto.Logger.WriteLineColorfully('    Text in clipboard is: <Color=Cyan>{}</Color>'.format(auto.GetClipboardText()))
        elif k == auto.ClipboardFormat.CF_HTML:
            htmlText = auto.GetClipboardHtml()
            auto.Logger.WriteLineColorfully('    Html text in clipboard is: <Color=Cyan>{}</Color>'.format(htmlText))
        elif k == auto.ClipboardFormat.CF_BITMAP:
            auto.Logger.WriteLineColorfully('    Bitmap clipboard is: <Color=Cyan>{}</Color>'.format(auto.GetClipboardBitmap()))

    auto.InputColorfully('paused, press Enter to test <Color=Cyan>SetClipboardText</Color>', auto.ConsoleColor.Green)
    auto.SetClipboardText('Hello World')
    auto.InputColorfully('<Color=Yellow>You can paste it in Office Word now.</Color>\npaused, press Enter to test <Color=Cyan>SetClipboardHtml</Color>')
    auto.SetClipboardHtml('<h1>Title</h1><br><h3>Hello</h3><br><p>test html</p><br>'),
    auto.InputColorfully('<Color=Yellow>You can paste it in Office Word now.</Color>\npaused, press Enter to test <Color=Cyan>SetClipboardBitmap</Color>')
    c = auto.ControlFromCursor()
    with c.ToBitmap() as bmp:
        auto.SetClipboardBitmap(bmp)
    auto.InputColorfully('<Color=Yellow>You can paste it in Office Word now.</Color>\npaused, press Enter to <Color=Cyan>exit</Color>')


if __name__ == '__main__':
    # testThread()
    main()
