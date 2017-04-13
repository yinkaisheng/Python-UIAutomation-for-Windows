# The uiautomation module

This module is for UIAutomatoin on Windows(Windows XP with SP3, Windows Vista, Windows 7 and Windows 8/8.1/10).
It supports UIAutomatoin for the applications which implmented IUIAutomation, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt and Firefox.

uiautomation is shared under the MIT Licence.
This means that the code can be freely copied and distributed, and costs nothing to use.

Only 3 files(**uiautomation.py, AutomationClientX86.dll and AutomationClientX64.dll**) are needed for UIAutomation. Other scripts are all demos.
You can install uiautomation by "pip install uiautomation"

Run 'uiautomation.py -h' for help.
Run automate_notepad_py3.py to see a simple demo.

Microsoft IUIAutomation Minimum supported client:
Windows 7, Windows Vista with SP2 and Platform Update for Windows Vista, Windows XP with SP3 and Platform Update for Windows Vista [desktop apps only]

Microsoft IUIAutomation Minimum supported server:
Windows Server 2008 R2, Windows Server 2008 with SP2 and Platform Update for Windows Server 2008, Windows Server 2003 with SP2 and Platform Update for Windows Server 2008 [desktop apps only]

If "RuntimeError: Can not get an instance of IUIAutomation" occured when running uiautomation.py,
You need to install update [KB971513](https://support.microsoft.com/en-us/kb/971513) for your Windows.
You can also download from here [https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation](https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation)

Another UI tool inspectX86.exe or inspectX64.exe supplied by Microsoft can also be used to see the UI elements.

[Inspect (Inspect.exe)](https://msdn.microsoft.com/en-us/library/windows/desktop/dd318521%28v=vs.85%29.aspx) is a Windows-based tool that enables you select any UI element and view the element's accessibility data. You can view Microsoft UI Automation properties and control patterns, as well as Microsoft Active Accessibility properties. Inspect also enables you to test the navigational structure of the automation elements in the UI Automation tree, and the accessible objects in the Microsoft Active Accessibility hierarchy.

Inspect is installed with the Windows Software Development Kit (SDK) for Windows 8. (It is also available in previous versions of Windows SDK.) It is located in the \bin\<platform> folder of the SDK installation path (Inspect.exe).

--------------------------------------------------------------------------------
How to use uiautomation?
run 'automation.py -h' or 'uiautomation.py -h'

run Notepad.exe, run cmd.exe and cd to uiautomation directory, run uiautomation.py in cmd and then swith to Notepad immediately
wait for 5 seconds
uiautomation will print the control tree of Notepad:

ControlType: WindowControl    ClassName: Notepad    AutomationId:     Rect: (372, 219, 1172, 669)    Name: 无标题 - 记事本    Handle: 0x70A62(461410)    Depth: 0
    ControlType: EditControl    ClassName: Edit    AutomationId: 15    Rect: (380, 269, 1164, 639)    Name:     Handle: 0x190A8E(1641102)    Depth: 1    Value: 
        ControlType: ScrollBarControl    ClassName:     AutomationId: NonClientVerticalScrollBar    Rect: (1145, 271, 1162, 620)    Name: 垂直滚动条    Handle: 0x0(0)    Depth: 2    RangeValue: 0
            ControlType: ButtonControl    ClassName:     AutomationId: UpButton    Rect: (1145, 271, 1162, 288)    Name: 上一行    Handle: 0x0(0)    Depth: 3
            ControlType: ButtonControl    ClassName:     AutomationId: DownButton    Rect: (1145, 603, 1162, 620)    Name: 下一行    Handle: 0x0(0)    Depth: 3
        ControlType: ScrollBarControl    ClassName:     AutomationId: NonClientHorizontalScrollBar    Rect: (382, 620, 1145, 637)    Name: 水平滚动条    Handle: 0x0(0)    Depth: 2    RangeValue: 0
            ControlType: ButtonControl    ClassName:     AutomationId: UpButton    Rect: (382, 620, 399, 637)    Name: 左移一列    Handle: 0x0(0)    Depth: 3
            ControlType: ButtonControl    ClassName:     AutomationId: DownButton    Rect: (1128, 620, 1145, 637)    Name: 右移一列    Handle: 0x0(0)    Depth: 3
        ControlType: ThumbControl    ClassName:     AutomationId:     Rect: (1145, 620, 1162, 637)    Name:     Handle: 0x0(0)    Depth: 2
    ControlType: StatusBarControl    ClassName: msctls_statusbar32    AutomationId: 1025    Rect: (380, 639, 1164, 661)    Name:     Handle: 0x150A6A(1378922)    Depth: 1
        ControlType: TextControl    ClassName:     AutomationId:     Rect: (380, 641, 968, 661)    Name:     Handle: 0x0(0)    Depth: 2
        ControlType: TextControl    ClassName:     AutomationId:     Rect: (970, 641, 1148, 661)    Name:    第 1 行，第 1 列    Handle: 0x0(0)    Depth: 2
    ControlType: TitleBarControl    ClassName:     AutomationId:     Rect: (396, 222, 1164, 249)    Name:     Handle: 0x0(0)    Depth: 1
        ControlType: MenuBarControl    ClassName:     AutomationId: MenuBar    Rect: (380, 227, 401, 248)    Name: 系统    Handle: 0x0(0)    Depth: 2
            ControlType: MenuItemControl    ClassName:     AutomationId:     Rect: (380, 227, 401, 248)    Name: 系统    Handle: 0x0(0)    Depth: 3    CurrentExpandCollapseState: ExpandCollapseState.Collapsed
        ControlType: ButtonControl    ClassName:     AutomationId:     Rect: (1061, 220, 1090, 240)    Name: 最小化    Handle: 0x0(0)    Depth: 2
        ControlType: ButtonControl    ClassName:     AutomationId:     Rect: (1090, 220, 1117, 240)    Name: 最大化    Handle: 0x0(0)    Depth: 2
        ControlType: ButtonControl    ClassName:     AutomationId:     Rect: (1117, 220, 1166, 240)    Name: 关闭    Handle: 0x0(0)    Depth: 2
    ControlType: MenuBarControl    ClassName:     AutomationId: MenuBar    Rect: (380, 249, 1164, 268)    Name: 应用程序    Handle: 0x0(0)    Depth: 1
        ControlType: MenuItemControl    ClassName:     AutomationId:     Rect: (380, 249, 432, 268)    Name: 文件(F)    Handle: 0x0(0)    Depth: 2    CurrentExpandCollapseState: ExpandCollapseState.Collapsed
        ControlType: MenuItemControl    ClassName:     AutomationId:     Rect: (432, 249, 485, 268)    Name: 编辑(E)    Handle: 0x0(0)    Depth: 2    CurrentExpandCollapseState: ExpandCollapseState.Collapsed
        ControlType: MenuItemControl    ClassName:     AutomationId:     Rect: (485, 249, 541, 268)    Name: 格式(O)    Handle: 0x0(0)    Depth: 2    CurrentExpandCollapseState: ExpandCollapseState.Collapsed
        ControlType: MenuItemControl    ClassName:     AutomationId:     Rect: (541, 249, 595, 268)    Name: 查看(V)    Handle: 0x0(0)    Depth: 2    CurrentExpandCollapseState: ExpandCollapseState.Collapsed
        ControlType: MenuItemControl    ClassName:     AutomationId:     Rect: (595, 249, 650, 268)    Name: 帮助(H)    Handle: 0x0(0)    Depth: 2    CurrentExpandCollapseState: ExpandCollapseState.Collapsed
```
import subprocess
import uiautomation as automation

subprocess.Popen('notepad.exe')
notepadWindow = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad')
edit = notepadWindow.EditControl()
edit.SetValue('Hello')
edit.SendKeys('{Ctrl}{End}{Enter}Hello')
```


[具体用法参考](http://www.cnblogs.com/Yinkaisheng/p/3444132.html)

Inspect
![Inspect](https://i-msdn.sec.s-msft.com/dynimg/IC510569.png)

WindowsDesktop
![Desktop](https://raw.githubusercontent.com/yinkaisheng/Python-UIAutomation-for-Windows/master/automation_desktop.png)

Qt5
![Qt5](https://raw.githubusercontent.com/yinkaisheng/Python-UIAutomation-for-Windows/master/automation_Qt.png)

Firefox
![Firefox](https://raw.githubusercontent.com/yinkaisheng/Python-UIAutomation-for-Windows/master/automation_firefox.png)

Wireshark(version must >= 2.0)
![Wireshark](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows/raw/master/wireshark_rtp_analyzer.png)

QQ
![QQ](https://raw.githubusercontent.com/yinkaisheng/Python-UIAutomation-for-Windows/master/automation_qq.png)

Batch rename pdf bookmark
![bookmark](https://raw.githubusercontent.com/yinkaisheng/Python-UIAutomation-for-Windows/master/rename_pdf_bookmark.gif)