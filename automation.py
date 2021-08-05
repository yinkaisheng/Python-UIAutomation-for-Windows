#!python3
# -*- coding:utf-8 -*-
import sys
import time

import uiautomation as auto


def usage():
    auto.Logger.ColorfullyWrite("""usage
<Color=Cyan>-h</Color>      show command <Color=Cyan>help</Color>
<Color=Cyan>-t</Color>      delay <Color=Cyan>time</Color>, default 3 seconds, begin to enumerate after Value seconds, this must be an integer
        you can delay a few seconds and make a window active so automation can enumerate the active window
<Color=Cyan>-d</Color>      enumerate tree <Color=Cyan>depth</Color>, this must be an integer, if it is null, enumerate the whole tree
<Color=Cyan>-r</Color>      enumerate from <Color=Cyan>root</Color>:Desktop window, if it is null, enumerate from foreground window
<Color=Cyan>-f</Color>      enumerate from <Color=Cyan>focused</Color> control, if it is null, enumerate from foreground window
<Color=Cyan>-c</Color>      enumerate the control under <Color=Cyan>cursor</Color>, if depth is < 0, enumerate from its ancestor up to depth
<Color=Cyan>-a</Color>      show <Color=Cyan>ancestors</Color> of the control under cursor
<Color=Cyan>-n</Color>      show control full <Color=Cyan>name</Color>, if it is null, show first 30 characters of control's name in console,
        always show full name in log file @AutomationLog.txt
<Color=Cyan>-p</Color>      show <Color=Cyan>process id</Color> of controls

if <Color=Red>UnicodeError</Color> or <Color=Red>LookupError</Color> occurred when printing,
try to change the active code page of console window by using <Color=Cyan>chcp</Color> or see the log file <Color=Cyan>@AutomationLog.txt</Color>
chcp, get current active code page
chcp 936, set active code page to gbk
chcp 65001, set active code page to utf-8

examples:
automation.py -t3
automation.py -t3 -r -d1 -m -n
automation.py -c -t3

""", writeToFile=False)


def main():
    import getopt
    auto.Logger.Write('UIAutomation {} (Python {}.{}.{}, {} bit)\n'.format(auto.VERSION, sys.version_info.major, sys.version_info.minor, sys.version_info.micro, 64 if sys.maxsize > 0xFFFFFFFF else 32))
    options, args = getopt.getopt(sys.argv[1:], 'hrfcanpd:t:',
                                  ['help', 'root', 'focus', 'cursor', 'ancestor', 'showAllName', 'depth=',
                                   'time='])
    root = False
    focus = False
    cursor = False
    ancestor = False
    foreground = True
    showAllName = False
    depth = 0xFFFFFFFF
    seconds = 3
    showPid = False
    for (o, v) in options:
        if o in ('-h', '-help'):
            usage()
            sys.exit(0)
        elif o in ('-r', '-root'):
            root = True
            foreground = False
        elif o in ('-f', '-focus'):
            focus = True
            foreground = False
        elif o in ('-c', '-cursor'):
            cursor = True
            foreground = False
        elif o in ('-a', '-ancestor'):
            ancestor = True
            foreground = False
        elif o in ('-n', '-showAllName'):
            showAllName = True
        elif o in ('-p', ):
            showPid = True
        elif o in ('-d', '-depth'):
            depth = int(v)
        elif o in ('-t', '-time'):
            seconds = int(v)
    if seconds > 0:
        auto.Logger.Write('please wait for {0} seconds\n\n'.format(seconds), writeToFile=False)
        time.sleep(seconds)
    auto.Logger.ColorfullyLog('Starts, Current Cursor Position: <Color=Cyan>{}</Color>'.format(auto.GetCursorPos()))
    control = None
    if root:
        control = auto.GetRootControl()
    if focus:
        control = auto.GetFocusedControl()
    if cursor:
        control = auto.ControlFromCursor()
        if depth < 0:
            while depth < 0 and control:
                control = control.GetParentControl()
                depth += 1
            depth = 0xFFFFFFFF
    if ancestor:
        control = auto.ControlFromCursor()
        if control:
            auto.EnumAndLogControlAncestors(control, showAllName, showPid)
        else:
            auto.Logger.Write('IUIAutomation returns null element under cursor\n', auto.ConsoleColor.Yellow)
    else:
        indent = 0
        if not control:
            control = auto.GetFocusedControl()
            controlList = []
            while control:
                controlList.insert(0, control)
                control = control.GetParentControl()
            if len(controlList) == 1:
                control = controlList[0]
            else:
                control = controlList[1]
                if foreground:
                    indent = 1
                    auto.LogControl(controlList[0], 0, showAllName, showPid)
        auto.EnumAndLogControl(control, depth, showAllName, showPid, startDepth=indent)
    auto.Logger.Log('Ends\n')


if __name__ == '__main__':
    main()
