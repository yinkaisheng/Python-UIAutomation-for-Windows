# -*- coding: utf-8 -*-
# this script only works with Win32 notepad.exe
# if you notepad.exe is the Windows Store version in Windows 11, you need to uninstall it.
import os, sys
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))  # not required after 'pip install uiautomation'
from uiautomation import uiautomation as auto

def test():
    # print(auto.GetRootControl())
    subprocess.Popen('notepad.exe', shell=True)
    # you should find the top level window first, then find children from the top level window
    notepadWindow = auto.WindowControl(searchDepth=1, ClassName='Notepad')
    if not notepadWindow.Exists(3, 1):
        print('Can not find Notepad window')
        exit(0)
    print(notepadWindow)
    notepadWindow.SetTopmost(True)
    # find the first EditControl in notepadWindow
    edit = notepadWindow.EditControl()
    # usually you don't need to catch exceptions
    # but if you meet a COMError exception, put it in a try block
    try:
        # use value pattern to get or set value
        edit.GetValuePattern().SetValue('Hello')# or edit.GetPattern(auto.PatternId.ValuePattern)
    except auto.comtypes.COMError as ex:
        # maybe you don't run python as administrator
        # or the control doesn't have a implementation for the pattern method(I have no solution for this)
        pass
    edit.Click() # this step is optional, but some edits may need it
    edit.SendKeys('{Ctrl}{End}{Enter}World')
    print('current text:', edit.GetValuePattern().Value)
    notepadWindow.CaptureToImage('notepad.png')
    notepadWindow.MenuBarControl(searchDepth=1).CaptureToImage('notepad_menubar.png')

    # generate an animated gif
    bitmap = notepadWindow.ToBitmap(x=0, y=0, width=160, height=160)
    side = int(bitmap.Width * 1.42)
    gifBmp = auto.Bitmap(side, side)
    gifBmp.Clear(0xFFFF_FFFF) # set bitmap background color to white
    gifBmp.Paste(x=(side-bitmap.Width)//2, y=(side-bitmap.Height)//2, bitmap=bitmap)
    gifFrameCount = 20
    bmps = [gifBmp.RotateWithSameSize(gifBmp.Width//2, gifBmp.Height//2, i*360/gifFrameCount) for i in range(0, gifFrameCount)]
    auto.GIF.ToGifFile('notepad_part.gif', bitmaps=bmps, delays=[100]*gifFrameCount)

    # find the first TitleBarControl in notepadWindow,
    # then find the second ButtonControl in TitleBarControl, which is the Maximize button
    maximizeButton = notepadWindow.TitleBarControl().ButtonControl(foundIndex=2)
    maximizeButton.Click(waitTime=2)
    maximizeButton.Click()
    # find the first button in notepadWindow whose Name is '关闭' or 'Close', the close button
    # the relative depth from Close button to Notepad window is 2
    notepadWindow.ButtonControl(searchDepth=2, Compare=lambda c, d: c.Name in ['Close', '关闭']).Click()
    # then notepad will popup a window askes you to save or not, press hotkey alt+n not to save
    notepadWindow.WindowControl(searchDepth=1).CaptureToImage('notepad_save.png')
    auto.SendKeys('{Alt}n')

if __name__ == '__main__':
    test()