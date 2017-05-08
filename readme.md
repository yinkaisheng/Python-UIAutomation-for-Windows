# The uiautomation module

This module is for UIAutomatoin on Windows(Windows XP with SP3, Windows Vista, Windows 7 and Windows 8/8.1/10).
It supports UIAutomatoin for the applications which implmented IUIAutomation, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt and Firefox.

uiautomation is shared under the Apache Licence 2.0.
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
You can also download it from here [https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation](https://github.com/yinkaisheng/WindowsUpdateKB971513ForIUIAutomation)

By the way, You'd better run python as administrator. Otherwise uiautomation may fail to enumerate controls under some circumstances.

--------------------------------------------------------------------------------
How to use uiautomation?
run '**automation.py -h**' or '**uiautomation.py -h**'
![help](https://github.com/yinkaisheng/Python-UIAutomation-for-Windows/raw/master/uiautomation-h.png)
Understand the arguments of uiautomation, and try the following examples  
**uiautomation.py -r -d 1 -t 0**, print desktop(the root of control tree) and it's children(top level windows)  
**uiautomation.py -t 0**, print current active window's controls  
  
run notepad.exe, run uiautomation.py -t 3, swith to Notepad and wait for 5 seconds  
  
uiautomation will print the controls of Notepad and save them to @AutomationLog.txt:  
  
ControlType: WindowControl    ClassName: Notepad  
　　ControlType: EditControl    ClassName: Edit  
　　　　ControlType: ScrollBarControl    ClassName:  
　　　　　　ControlType: ButtonControl    ClassName:  
　　　　　　ControlType: ButtonControl    ClassName:  
　　　　ControlType: ScrollBarControl    ClassName:  
　　　　　　ControlType: ButtonControl    ClassName:  
　　　　　　ControlType: ButtonControl    ClassName:  
　　　　ControlType: ThumbControl    ClassName:  
　　ControlType: StatusBarControl    ClassName:  
　　　　ControlType: TextControl    ClassName:  
　　　　ControlType: TextControl    ClassName:  
...  

run the following code
```python
import subprocess
import uiautomation as automation

print(automation.GetRootControl())
subprocess.Popen('notepad.exe')
notepadWindow = automation.WindowControl(searchDepth = 1, ClassName = 'Notepad')
print(notepadWindow.Name)
notepadWindow.SetTopmost(True)
edit = notepadWindow.EditControl()
edit.SetValue('Hello')
edit.SendKeys('{Ctrl}{End}{Enter}World')

```
automation.GetRootControl() returns the root control  
automation.WindowControl(searchDepth = 1, ClassName = 'Notepad') creates a WindowControl, the parameters specify how to search the control  
the following parameters can be used  
searchFromControl = None,   
searchDepth = 0xFFFFFFFF,   
searchWaitTime = SEARCH_INTERVAL,   
foundIndex = 1  
Name  
ClassName  
AutomationId  
ControlType  
Depth  

See Control.\_\_init\_\_ for the comment of the parameters  
See automation_notepad_py3.py for a detailed example  

--------------------------------------------------------------------------------
Another UI tool inspectX86.exe or inspectX64.exe supplied by Microsoft can also be used to see the UI elements.

[Inspect (Inspect.exe)](https://msdn.microsoft.com/en-us/library/windows/desktop/dd318521%28v=vs.85%29.aspx) is a Windows-based tool that enables you select any UI element and view the element's accessibility data. You can view Microsoft UI Automation properties and control patterns, as well as Microsoft Active Accessibility properties. Inspect also enables you to test the navigational structure of the automation elements in the UI Automation tree, and the accessible objects in the Microsoft Active Accessibility hierarchy.

Inspect is installed with the Windows Software Development Kit (SDK) for Windows 8. (It is also available in previous versions of Windows SDK.) It is located in the \bin\<platform> folder of the SDK installation path (Inspect.exe).

Inspect
![Inspect](https://i-msdn.sec.s-msft.com/dynimg/IC510569.png)
--------------------------------------------------------------------------------

[代码原理简单介绍](http://www.cnblogs.com/Yinkaisheng/p/3444132.html)

Some screenshots:

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