# The uiautomation module

This module is for [UIAutomation](https://docs.microsoft.com/en-us/windows/desktop/WinAuto/windows-automation-api-portal) on Windows(Windows XP with SP3, Windows Vista, Windows 7 and Windows 8/8.1/10).
It supports UIAutomation for the applications which implmented UIAutomation Provider, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt(Partly), Firefox(**version<=56 or >=60**) and Chrome(requires **--force-renderer-accessibility** command line parameter).

I developed it in my spare time and for my personal use.
I only tested it with official CPython 2.7, 3.4, 3.5 and 3.6.

uiautomation is shared under the Apache Licence 2.0.
This means that the code can be freely copied and distributed, and costs nothing to use.

You can install uiautomation by "pip install uiautomation". After installation, a automation.py that calls uiautomation will be in 'C:\PythonXX\Scripts\'.
You use this script to traverse UI controls.

Run 'C:\PythonXX\Scripts\automation.py -h' for help.
Run automate_notepad_py3.py to see a simple demo.

If "RuntimeError: Can not get an instance of IUIAutomation" occured when running uiautomation.py,
You need to install update [KB971513](https://support.microsoft.com/en-us/kb/971513) for your Windows(Mostly Windows XP).
You can also download it from here [https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation](https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation)

On Windows 8/8.1, to automate a Metro App, the app must be in foreground. If a Metro App was switched to background, uiautomation can't fetch its controls' information.

By the way, You'd better run python as administrator. Otherwise uiautomation may fail to enumerate controls under some circumstances.

[Requirements:](https://docs.microsoft.com/zh-cn/windows/desktop/api/uiautomationclient/nn-uiautomationclient-iuiautomation)

Microsoft UIAutomation Minimum supported client:
Windows 7, Windows Vista with SP2 and Platform Update for Windows Vista, Windows XP with SP3 and Platform Update for Windows Vista [desktop apps only]

Microsoft UIAutomation Minimum supported server:
Windows Server 2008 R2, Windows Server 2008 with SP2 and Platform Update for Windows Server 2008, Windows Server 2003 with SP2 and Platform Update for Windows Server 2008 [desktop apps only]

C++ dll source code: [UIAutomationClient](https://github.com/yinkaisheng/UIAutomationClient)

--------------------------------------------------------------------------------
How to use uiautomation?
run '**automation.py -h**'
![help](images/uiautomation-h.png)

Understand the arguments of automation.py, and try the following examples  
**automation.py -r -d 1 -t 0**, print desktop(the root of control tree) and it's children(top level windows)  
**automation.py -t 0 -n -m**, print current active window's controls, show fullname, show more properties
  
run notepad.exe, run automation.py -t 3, swith to Notepad and wait for 5 seconds  
  
automation.py will print the controls of Notepad and save them to @AutomationLog.txt:  
  
ControlType: PaneControl    ClassName: #32769    Name: 桌面    Depth: 0    **(Desktop window, the root control)**  
　　ControlType: WindowControl    ClassName: Notepad    Depth: 1    **(Top level window)**  
　　　　ControlType: EditControl    ClassName: Edit    Depth: 2  
　　　　　　ControlType: ScrollBarControl    ClassName:     Depth: 3  
　　　　　　　　ControlType: ButtonControl    ClassName:     Depth: 4  
　　　　　　　　ControlType: ButtonControl    ClassName:     Depth: 4  
　　　　　　ControlType: ThumbControl    ClassName:     Depth: 3  
　　　　ControlType: TitleBarControl    ClassName:     Depth: 2  
　　　　　　ControlType: MenuBarControl    ClassName:     Depth: 3  
　　　　　　　　ControlType: MenuItemControl    ClassName:     Depth: 4  
　　　　　　ControlType: ButtonControl    ClassName:     Name: 最小化    Depth: 3    **(Minimize Button)**  
　　　　　　ControlType: ButtonControl    ClassName:     Name: 最大化    Depth: 3    **(Maximize Button)**  
　　　　　　ControlType: ButtonControl    ClassName:     Name: 关闭    Depth: 3    **(Close Button)**  
...  

run the following code
```python
# -*- coding: utf-8 -*-
import subprocess
import uiautomation as auto

print(auto.GetRootControl())
subprocess.Popen('notepad.exe')
# you should find the top level window first, then find children from the top level window
notepadWindow = auto.WindowControl(searchDepth = 1, ClassName = 'Notepad')
print(notepadWindow.Name)
notepadWindow.SetTopmost(True)
# find the first EditControl in notepadWindow
edit = notepadWindow.EditControl()
edit.SetValue('Hello')
edit.SendKeys('{Ctrl}{End}{Enter}World')
# find the first TitleBarControl in notepadWindow, 
# then find the second ButtonControl in TitleBarControl, which is the Maximize button
notepadWindow.TitleBarControl().ButtonControl(foundIndex = 2).Click()
# find the first button in notepadWindow whose Name is '关闭', the close button, the relative depth from Close button to Notepad window is 2
notepadWindow.ButtonControl(searchDepth = 2, Name = '关闭').Click()
# then notepad will popup a window askes you to save or not, press hotkey alt+n not to save
auto.SendKeys('{Alt}n')
```
automation.GetRootControl() returns the root control(the Desktop window)  
automation.WindowControl(searchDepth = 1, ClassName = 'Notepad') creates a WindowControl, the parameters specify how to search the control  
the following parameters can be used  
searchFromControl = None,   
searchDepth = 0xFFFFFFFF,   
searchWaitTime = SEARCH_INTERVAL,   
foundIndex = 1  
Name  
SubName  
RegexName  
ClassName  
AutomationId  
ControlType  
Depth  
Compare  

See Control.\_\_init\_\_ for the comments of the parameters  
See scripts in folder **demos** for more examples  

**If automation.py can't print the controls you see.
Maybe the controls were built by DirectUI(or CustomControl), not UI Frameworks supplied by Microsoft.
In order to support UIAutomation, an UI Framework must implement [UI Automation Provider](https://docs.microsoft.com/zh-cn/windows/desktop/WinAuto/uiauto-providersoverview).**

A Microsoft UI Automation provider is a software object that exposes an element of an application's UI so that accessibility client applications can retrieve information about the element and invoke its functionality. In general, each control or other distinct element in a UI has a provider.

Microsoft includes a provider for each of the standard controls that are supplied with Microsoft Win32, Windows Forms, and Windows Presentation Foundation (WPF). This means that the standard controls are automatically exposed to UI Automation clients; you do not need to implement any accessibility interfaces for the standard controls.

If your application includes any custom controls, you need to implement UI Automation providers for those controls to make them accessible to accessibility client applications. You also need to implement providers for any third party controls that do not include a provider. You implement a provider by implementing UI Automation provider interfaces and control pattern interfaces.

--------------------------------------------------------------------------------
Another UI tool [Inspect.exe](https://docs.microsoft.com/zh-cn/windows/desktop/WinAuto/inspect-objects) supplied by Microsoft can also be used to traverse the UI elements. It has an UI interface while my script shows UI elements in terminal.
But I found that my script is more convenient sometimes.

![Inspect](https://docs.microsoft.com/zh-cn/windows/desktop/WinAuto/images/inspect.png)

--------------------------------------------------------------------------------

[代码原理简单介绍(Chinese)](http://www.cnblogs.com/Yinkaisheng/p/3444132.html)

--------------------------------------------------------------------------------
Some screenshots:

Batch rename pdf bookmark
![bookmark](images/rename_pdf_bookmark.gif)

Firefox
![Firefox](images/automation_firefox.png)

WindowsDesktop
![Desktop](images/automation_desktop.png)

Qt5
![Qt5](images/automation_Qt.png)

Wireshark(version must >= 2.0)
![Wireshark](images/wireshark_rtp_analyzer.png)

QQ
![QQ](images/automation_qq.png)

Pretty print dir
![PrettyPrint](images/pretty_print_dir.png)