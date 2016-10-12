# The automation module

This module is for automation on Windows(Windows XP with SP3, Windows Vista, Windows 7 and Windows 8/8.1/10).
It supports automation for the applications which implmented IUIAutomation, such as MFC, Windows Form, WPF, Modern UI(Metro UI), Qt and Firefox.

automation is shared under the MIT Licence.
This means that the code can be freely copied and distributed, and costs nothing to use.

Only 3 files(<font color=blue>**automation.py, AutomationClientX86.dll and AutomationClientX64.dll**</font>) are needed for UIAutomation. Other scripts are all demos.

Run 'automation.py -h' for help.
Run automate_notepad_py3.py to see a simple demo.

Another UI tool inspectX86.exe or inspectX64.exe supplied by Microsoft can also be used to see the UI elements.

[Inspect (Inspect.exe)](https://msdn.microsoft.com/en-us/library/windows/desktop/dd318521%28v=vs.85%29.aspx) is a Windows-based tool that enables you select any UI element and view the element's accessibility data. You can view Microsoft UI Automation properties and control patterns, as well as Microsoft Active Accessibility properties. Inspect also enables you to test the navigational structure of the automation elements in the UI Automation tree, and the accessible objects in the Microsoft Active Accessibility hierarchy.

Inspect is installed with the Windows Software Development Kit (SDK) for Windows 8. (It is also available in previous versions of Windows SDK.) It is located in the \bin\<platform> folder of the SDK installation path (Inspect.exe).

--------------------------------------------------------------------------------

Author mail: yinkaisheng@foxmail.com

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